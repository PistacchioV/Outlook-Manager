# -*- coding: utf-8 -*-
"""
outlook_manager.py
==================

Camada de domínio do "Gerenciador de E-mails Inteligente".

Responsabilidades:
    * Conexão com o Outlook local via ``win32com.client`` (pywin32).
    * Varredura periódica do Inbox em uma *background thread*.
    * Filtragem por palavras-chave OU remetentes-chave.
    * Agrupamento de e-mails por tópico (assunto normalizado / ConversationID).
    * Geração de um resumo curto por tópico (placeholder plugável a um LLM).

IMPORTANTE — Concorrência COM no Windows:
    O Outlook é exposto via COM (Component Object Model). Toda thread que
    fala com o COM precisa inicializá-lo *naquela thread* com
    ``pythoncom.CoInitialize()`` e liberá-lo no fim com
    ``pythoncom.CoUninitialize()``. Como a varredura roda numa thread
    separada da do Flask, isso é feito dentro de ``_worker_loop``.

IMPORTANTE — Plataforma:
    ``pywin32`` só existe no Windows com Outlook desktop instalado.
    Em outros sistemas (ou se o import falhar) o módulo entra em
    ``MODO_SIMULADO``, servindo dados fictícios para que a interface web
    possa ser desenvolvida/testada em qualquer máquina.
"""

from __future__ import annotations

import json
import os
import re
import threading
import time
import unicodedata
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

# Arquivo de persistência da configuração (mesmo diretório do módulo).
_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

# ---------------------------------------------------------------------------
# Import condicional do pywin32. Fora do Windows caímos em MODO_SIMULADO.
# ---------------------------------------------------------------------------
try:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    MODO_SIMULADO = False
except Exception:  # ImportError no Mac/Linux, ou ambiente sem Outlook.
    pythoncom = None  # type: ignore
    win32com = None  # type: ignore
    MODO_SIMULADO = True


# ===========================================================================
# RESUMO AUTOMÁTICO (placeholder)
# ===========================================================================
def gerar_resumo(texto: str, max_frases: int = 3) -> str:
    """Gera um resumo curto (2–3 frases) de um corpo de e-mail.

    Implementação **placeholder** baseada em heurística simples: limpa o
    texto e devolve as primeiras ``max_frases`` frases relevantes.

    >>> PONTO DE EXTENSÃO PARA LLM <<<
    Para usar uma LLM de verdade (ex.: Claude), substitua o corpo desta
    função por uma chamada de API. Exemplo com o SDK da Anthropic::

        from anthropic import Anthropic
        client = Anthropic(api_key="...")

        def gerar_resumo(texto, max_frases=3):
            msg = client.messages.create(
                model="claude-opus-4-8",
                max_tokens=200,
                messages=[{
                    "role": "user",
                    "content": (
                        "Resuma o e-mail abaixo em no máximo 3 frases, "
                        "em português, foco no que requer ação:\\n\\n" + texto
                    ),
                }],
            )
            return msg.content[0].text.strip()

    Args:
        texto: Corpo do e-mail (texto puro).
        max_frases: Número máximo de frases no resumo.

    Returns:
        Uma string com o resumo. Nunca lança exceção.
    """
    if not texto:
        return "Sem conteúdo para resumir."

    # Normaliza espaços/quebras de linha e remove assinaturas óbvias.
    limpo = re.sub(r"\s+", " ", texto).strip()
    # Corta cadeias de resposta antigas ("De:", "From:", "-----").
    limpo = re.split(r"(?:^|\s)(?:De:|From:|-{3,}|_{3,})", limpo)[0].strip()

    # Quebra em frases por pontuação terminal.
    frases = re.split(r"(?<=[.!?])\s+", limpo)
    frases = [f.strip() for f in frases if len(f.strip()) > 0]

    if not frases:
        return limpo[:200] + ("..." if len(limpo) > 200 else "")

    resumo = " ".join(frases[:max_frases])
    # Garante que não fique gigante mesmo com frases longas.
    if len(resumo) > 320:
        resumo = resumo[:317].rstrip() + "..."
    return resumo


# ===========================================================================
# UTILITÁRIOS
# ===========================================================================
def _normalizar(texto: str) -> str:
    """Minúsculas, sem acentos e sem espaços nas pontas (p/ comparações)."""
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))
    return sem_acento.lower().strip()


def _chave_topico(assunto: str) -> str:
    """Normaliza o assunto removendo prefixos de resposta/encaminhamento.

    Faz com que "RE: Bug no login", "FW: bug no login" e "Bug no Login"
    caiam todos no mesmo tópico.
    """
    base = _normalizar(assunto)
    # Remove repetidamente prefixos tipo "re:", "res:", "fw:", "fwd:", "enc:".
    while True:
        novo = re.sub(r"^(re|res|fw|fwd|enc|encaminhar)\s*:\s*", "", base)
        if novo == base:
            break
        base = novo
    return base or "(sem assunto)"


# ===========================================================================
# GERENCIADOR PRINCIPAL
# ===========================================================================
class OutlookManager:
    """Mantém configuração, executa o worker e guarda os tópicos encontrados.

    Thread-safe: todo acesso ao estado compartilhado (config + tópicos) é
    protegido por um ``threading.Lock``.
    """

    def __init__(self, intervalo_segundos: int = 300) -> None:
        # ----- Configuração editável pela UI -----
        self.palavras_chave: List[str] = ["urgente", "faturamento", "bug"]
        self.pessoas_chave: List[str] = []
        self.intervalo_segundos: int = intervalo_segundos
        # Conta/mailbox do Outlook a ser lida. Numa máquina corporativa pode
        # haver várias contas; aqui miramos uma SMTP específica. Vazio = usa a
        # conta padrão do Outlook (GetDefaultFolder).
        self.conta_email: str = "giulliano.luccia@jpmorgan.com"

        # ----- Estado de resultados -----
        # Dicionário: chave_topico -> dados do tópico (ver _montar_topico).
        self._topicos: Dict[str, Dict[str, Any]] = {}
        self.ultima_varredura: Optional[str] = None
        self.status_worker: str = "parado"
        # Conta que o Outlook REALMENTE reconheceu na última varredura
        # bem-sucedida (None enquanto não conectou / em erro).
        self.conta_conectada: Optional[str] = None

        # ----- Concorrência -----
        self._lock = threading.Lock()
        self._stop_event = threading.Event()       # sinaliza parada do worker
        self._wake_event = threading.Event()       # sinaliza "varrer agora"
        self._thread: Optional[threading.Thread] = None

        # ----- Persistência -----
        # Carrega config.json se existir (sobrescreve os defaults acima).
        self._carregar_config()

    # ----------------------------------------------------------------- #
    # Persistência da configuração                                       #
    # ----------------------------------------------------------------- #
    def _carregar_config(self) -> None:
        """Lê config.json (se houver) para restaurar palavras/pessoas/intervalo."""
        try:
            with open(_CONFIG_PATH, "r", encoding="utf-8") as fp:
                dados = json.load(fp)
            self.palavras_chave = list(dados.get("palavras_chave", self.palavras_chave))
            self.pessoas_chave = list(dados.get("pessoas_chave", self.pessoas_chave))
            self.intervalo_segundos = int(
                dados.get("intervalo_segundos", self.intervalo_segundos)
            )
            self.conta_email = str(dados.get("conta_email", self.conta_email))
        except FileNotFoundError:
            pass  # primeira execução: mantém os defaults
        except Exception:
            pass  # config corrompido não deve impedir a inicialização

    def _salvar_config(self) -> None:
        """Grava a configuração atual em config.json (escrita atômica).

        Pré-condição: chamado com ``self._lock`` já adquirido.
        """
        dados = {
            "palavras_chave": self.palavras_chave,
            "pessoas_chave": self.pessoas_chave,
            "intervalo_segundos": self.intervalo_segundos,
            "conta_email": self.conta_email,
        }
        try:
            tmp = _CONFIG_PATH + ".tmp"
            with open(tmp, "w", encoding="utf-8") as fp:
                json.dump(dados, fp, ensure_ascii=False, indent=2)
            os.replace(tmp, _CONFIG_PATH)  # troca atômica (evita arquivo parcial)
        except Exception:
            pass  # falha ao persistir não deve derrubar a aplicação

    # ----------------------------------------------------------------- #
    # API de configuração (chamada pelas rotas Flask)                    #
    # ----------------------------------------------------------------- #
    def add_palavra(self, palavra: str) -> None:
        palavra = palavra.strip()
        if not palavra:
            return
        with self._lock:
            if _normalizar(palavra) not in {_normalizar(p) for p in self.palavras_chave}:
                self.palavras_chave.append(palavra)
            self._salvar_config()

    def remove_palavra(self, palavra: str) -> None:
        with self._lock:
            self.palavras_chave = [
                p for p in self.palavras_chave
                if _normalizar(p) != _normalizar(palavra)
            ]
            self._salvar_config()

    def add_pessoa(self, email: str) -> None:
        email = email.strip()
        if not email:
            return
        with self._lock:
            if _normalizar(email) not in {_normalizar(p) for p in self.pessoas_chave}:
                self.pessoas_chave.append(email)
            self._salvar_config()

    def remove_pessoa(self, email: str) -> None:
        with self._lock:
            self.pessoas_chave = [
                p for p in self.pessoas_chave
                if _normalizar(p) != _normalizar(email)
            ]
            self._salvar_config()

    def set_intervalo(self, segundos: int) -> None:
        with self._lock:
            self.intervalo_segundos = max(30, int(segundos))
            self._salvar_config()

    def set_conta(self, email: str) -> None:
        """Define a conta/mailbox alvo. Vazio = conta padrão do Outlook."""
        with self._lock:
            self.conta_email = email.strip()
            self._salvar_config()
        # Força nova varredura para refletir a troca de caixa na hora.
        self._wake_event.set()

    def get_config(self) -> Dict[str, Any]:
        with self._lock:
            return {
                "palavras_chave": list(self.palavras_chave),
                "pessoas_chave": list(self.pessoas_chave),
                "intervalo_segundos": self.intervalo_segundos,
                "conta_email": self.conta_email,
            }

    def get_topicos(self) -> Dict[str, Any]:
        """Retorna os tópicos ordenados do mais recente para o mais antigo."""
        with self._lock:
            topicos = sorted(
                self._topicos.values(),
                key=lambda t: t["ultima_atualizacao_ord"],
                reverse=True,
            )
            return {
                "topicos": topicos,
                "ultima_varredura": self.ultima_varredura,
                "status_worker": self.status_worker,
                "modo_simulado": MODO_SIMULADO,
                "conta_email": self.conta_email,        # conta configurada
                "conta_conectada": self.conta_conectada,  # conta reconhecida
            }

    # ----------------------------------------------------------------- #
    # Ciclo de vida do worker                                           #
    # ----------------------------------------------------------------- #
    def iniciar_worker(self) -> None:
        """Sobe a thread de varredura (idempotente)."""
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self._thread = threading.Thread(
            target=self._worker_loop, name="OutlookWorker", daemon=True
        )
        self._thread.start()

    def parar_worker(self) -> None:
        """Sinaliza parada e aguarda a thread encerrar."""
        self._stop_event.set()
        if self._thread:
            self._thread.join(timeout=10)

    def forcar_varredura(self) -> None:
        """Interrompe a espera atual para varrer imediatamente."""
        # O loop do worker dorme em fatias dentro de ``_wake_event.wait``;
        # setar este evento o faz sair da espera e varrer na hora.
        self._wake_event.set()

    # ----------------------------------------------------------------- #
    # Loop principal da thread                                          #
    # ----------------------------------------------------------------- #
    def _worker_loop(self) -> None:
        """Loop executado na background thread.

        Inicializa o COM **uma vez** para esta thread (obrigatório no
        Windows) e o libera ao sair. Entre varreduras, dorme em fatias
        para responder rápido a pedidos de parada / varredura forçada.
        """
        if not MODO_SIMULADO:
            pythoncom.CoInitialize()  # <-- obrigatório por thread (COM)

        self.status_worker = "rodando"
        try:
            # Primeira varredura imediata ao iniciar.
            self._executar_varredura_segura()

            while not self._stop_event.is_set():
                with self._lock:
                    intervalo = self.intervalo_segundos

                # Dorme em fatias de 1s para reagir a stop/wake rapidamente.
                dormiu = 0
                while dormiu < intervalo and not self._stop_event.is_set():
                    if self._wake_event.wait(timeout=1):
                        self._wake_event.clear()
                        break
                    dormiu += 1

                if self._stop_event.is_set():
                    break

                self._executar_varredura_segura()
        finally:
            self.status_worker = "parado"
            if not MODO_SIMULADO:
                pythoncom.CoUninitialize()  # <-- libera o COM da thread

    def _executar_varredura_segura(self) -> None:
        """Envolve a varredura em try/except para o loop nunca morrer."""
        try:
            if MODO_SIMULADO:
                emails = self._coletar_emails_simulados()
                # No simulado, "confirmamos" a conta configurada.
                self.conta_conectada = self.conta_email or "conta padrão (simulado)"
            else:
                emails = self._coletar_emails_outlook()
            self._processar_emails(emails)
            self.ultima_varredura = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        except Exception as exc:  # noqa: BLE001  (logamos e seguimos)
            self.status_worker = f"erro: {exc}"
            self.conta_conectada = None  # não confirmamos a caixa

    # ----------------------------------------------------------------- #
    # Coleta — Outlook real                                             #
    # ----------------------------------------------------------------- #
    def _coletar_emails_outlook(self) -> List[Dict[str, Any]]:
        """Lê o Inbox do Outlook e devolve uma lista de dicts simples.

        Retorna apenas e-mails recentes (últimos 7 dias) para limitar custo.
        """
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
        except Exception as exc:
            # Causa nº1 em ambiente corporativo: o "Novo Outlook" (rollout via
            # GPO) NÃO expõe automação COM — só o "Outlook Clássico" expõe.
            # Também cai aqui se o Outlook não estiver instalado/aberto.
            raise RuntimeError(
                "Não foi possível conectar ao Outlook via COM. "
                "Verifique se o OUTLOOK CLÁSSICO está instalado (o 'Novo Outlook' "
                "não suporta COM). No app, desligue o botão 'Novo Outlook'. "
                f"Detalhe técnico: {exc}"
            ) from exc

        # Seleciona a Inbox da conta-alvo (ou a padrão se não configurada).
        inbox = self._obter_inbox(namespace)

        itens = inbox.Items
        itens.Sort("[ReceivedTime]", True)  # mais novos primeiro

        # Filtro DASL/Restrict por data para não varrer a caixa inteira.
        limite = datetime.now() - timedelta(days=7)
        filtro = "[ReceivedTime] >= '" + limite.strftime("%m/%d/%Y %H:%M %p") + "'"
        try:
            itens = itens.Restrict(filtro)
        except Exception:
            pass  # se o filtro falhar, seguimos com todos os itens

        emails: List[Dict[str, Any]] = []
        for item in itens:
            try:
                # 43 == olMail (Class do MailItem). Ignora convites, etc.
                if getattr(item, "Class", 43) != 43:
                    continue

                remetente = self._extrair_email_remetente(item)
                emails.append({
                    "assunto": getattr(item, "Subject", "") or "(sem assunto)",
                    "remetente": remetente,
                    "remetente_nome": getattr(item, "SenderName", "") or remetente,
                    "corpo": getattr(item, "Body", "") or "",
                    "recebido": self._to_datetime(getattr(item, "ReceivedTime", None)),
                    "conversation_id": getattr(item, "ConversationID", "") or "",
                })
            except Exception:
                continue  # item problemático não derruba a varredura
        return emails

    def _obter_inbox(self, namespace: Any) -> Any:
        """Retorna a Inbox da conta-alvo (``self.conta_email``).

        Estratégia (numa máquina corporativa há várias contas/stores):
          1. Procura a Account cujo SMTP bate com ``conta_email`` e usa o
             Inbox do seu ``DeliveryStore`` (independe do idioma do Outlook).
          2. Fallback: procura um Store cujo nome contenha o e-mail.
          3. Fallback final: ``GetDefaultFolder(6)`` (conta padrão).

        Levanta ``RuntimeError`` se a conta foi configurada mas não foi
        encontrada, para o erro aparecer no painel.
        """
        with self._lock:
            alvo = _normalizar(self.conta_email)

        # 6 == olFolderInbox (constante padrão do Outlook).
        if not alvo:
            inbox = namespace.GetDefaultFolder(6)
            # Registra o nome real da store conectada (conta padrão).
            try:
                self.conta_conectada = inbox.Store.DisplayName or "conta padrão"
            except Exception:
                self.conta_conectada = "conta padrão"
            return inbox

        # 1) Por Account (forma mais robusta).
        try:
            for account in namespace.Session.Accounts:
                smtp = _normalizar(getattr(account, "SmtpAddress", "") or "")
                if smtp == alvo:
                    store = account.DeliveryStore
                    self.conta_conectada = account.SmtpAddress  # conta confirmada
                    return store.GetDefaultFolder(6)
        except Exception:
            pass

        # 2) Por Store (nome costuma ser o próprio e-mail).
        try:
            for store in namespace.Stores:
                nome = _normalizar(getattr(store, "DisplayName", "") or "")
                if alvo in nome:
                    self.conta_conectada = store.DisplayName  # conta confirmada
                    return store.GetDefaultFolder(6)
        except Exception:
            pass

        # 3) Não achou a conta configurada: erro explícito no painel.
        raise RuntimeError(
            f"Conta '{self.conta_email}' não encontrada no Outlook. "
            "Verifique se o perfil está logado nessa caixa (Arquivo > "
            "Configurações de Conta) ou ajuste o e-mail no painel."
        )

    @staticmethod
    def _extrair_email_remetente(item: Any) -> str:
        """Obtém o SMTP do remetente, lidando com contas Exchange."""
        # Em contas Exchange o SenderEmailAddress vem como /O=.../CN=...
        # então tentamos o PropertyAccessor para pegar o SMTP real.
        try:
            tipo = getattr(item, "SenderEmailType", "")
            if tipo == "EX":
                PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                sender = item.Sender
                if sender is not None:
                    exch = sender.GetExchangeUser()
                    if exch is not None:
                        return exch.PrimarySmtpAddress or ""
                    return sender.PropertyAccessor.GetProperty(PR_SMTP) or ""
        except Exception:
            pass
        return getattr(item, "SenderEmailAddress", "") or ""

    @staticmethod
    def _to_datetime(valor: Any) -> datetime:
        """Converte ReceivedTime (pywintypes.datetime) em datetime nativo."""
        if valor is None:
            return datetime.now()
        try:
            return datetime(
                valor.year, valor.month, valor.day,
                valor.hour, valor.minute, valor.second,
            )
        except Exception:
            return datetime.now()

    # ----------------------------------------------------------------- #
    # Coleta — modo simulado (fora do Windows)                          #
    # ----------------------------------------------------------------- #
    def _coletar_emails_simulados(self) -> List[Dict[str, Any]]:
        """Gera e-mails fictícios para desenvolver a UI sem Outlook."""
        agora = datetime.now()
        return [
            {
                "assunto": "RE: Bug no login do portal",
                "remetente": "maria.silva@empresa.com",
                "remetente_nome": "Maria Silva",
                "corpo": (
                    "Pessoal, o bug no login voltou a acontecer em produção. "
                    "Usuários não conseguem autenticar desde as 14h. "
                    "É urgente, precisamos de um hotfix hoje ainda."
                ),
                "recebido": agora - timedelta(minutes=4),
                "conversation_id": "CONV-BUG-LOGIN",
            },
            {
                "assunto": "Fwd: Faturamento do mês de junho",
                "remetente": "financeiro@empresa.com",
                "remetente_nome": "Setor Financeiro",
                "corpo": (
                    "Segue em anexo o relatório de faturamento de junho. "
                    "Favor revisar os valores destacados e confirmar até sexta."
                ),
                "recebido": agora - timedelta(minutes=18),
                "conversation_id": "CONV-FATURAMENTO",
            },
            {
                "assunto": "Bug no login do portal",
                "remetente": "joao.dev@empresa.com",
                "remetente_nome": "João Dev",
                "corpo": (
                    "Identifiquei a causa: o token expira cedo demais. "
                    "Estou subindo um patch agora para o ambiente de staging."
                ),
                "recebido": agora - timedelta(minutes=2),
                "conversation_id": "CONV-BUG-LOGIN",
            },
            {
                "assunto": "Reunião de alinhamento semanal",
                "remetente": "rh@empresa.com",
                "remetente_nome": "RH",
                "corpo": "Lembrete da reunião de quinta às 10h. Sem urgência.",
                "recebido": agora - timedelta(hours=3),
                "conversation_id": "CONV-RH",
            },
        ]

    # ----------------------------------------------------------------- #
    # Processamento: filtra, agrupa e resume                            #
    # ----------------------------------------------------------------- #
    def _email_relevante(self, email: Dict[str, Any]) -> Optional[str]:
        """Decide se o e-mail interessa. Retorna o motivo, ou None.

        Critério: palavra-chave no assunto/corpo OU remetente cadastrado.
        """
        with self._lock:
            palavras = [_normalizar(p) for p in self.palavras_chave]
            pessoas = [_normalizar(p) for p in self.pessoas_chave]

        remetente = _normalizar(email["remetente"])
        if remetente and remetente in pessoas:
            return f"Remetente-chave: {email['remetente']}"

        texto = _normalizar(email["assunto"] + " " + email["corpo"])
        for palavra in palavras:
            if palavra and palavra in texto:
                return f"Palavra-chave: {palavra}"

        return None

    def _processar_emails(self, emails: List[Dict[str, Any]]) -> None:
        """Filtra os e-mails relevantes e os agrupa em tópicos."""
        novos_topicos: Dict[str, Dict[str, Any]] = {}

        for email in emails:
            motivo = self._email_relevante(email)
            if motivo is None:
                continue

            # Chave de agrupamento: ConversationID (se houver) tem
            # prioridade; senão, assunto normalizado.
            chave = email.get("conversation_id") or _chave_topico(email["assunto"])

            topico = novos_topicos.get(chave)
            if topico is None:
                topico = {
                    "chave": chave,
                    "assunto": re.sub(
                        r"^(RE|RES|FW|FWD|ENC)\s*:\s*", "",
                        email["assunto"], flags=re.IGNORECASE,
                    ).strip() or "(sem assunto)",
                    "mensagens": [],
                    "pessoas": set(),
                    "motivos": set(),
                }
                novos_topicos[chave] = topico

            topico["mensagens"].append(email)
            topico["pessoas"].add(email["remetente_nome"] or email["remetente"])
            topico["motivos"].add(motivo)

        # Monta a estrutura final serializável (resumo + metadados).
        finais: Dict[str, Dict[str, Any]] = {}
        for chave, topico in novos_topicos.items():
            finais[chave] = self._montar_topico(topico)

        with self._lock:
            self._topicos = finais

    def _montar_topico(self, topico: Dict[str, Any]) -> Dict[str, Any]:
        """Constrói o dict final de um tópico (pronto para virar JSON)."""
        mensagens = sorted(topico["mensagens"], key=lambda m: m["recebido"])
        ultima = mensagens[-1]

        # Resumo do tópico: concatena os corpos e resume (placeholder/LLM).
        corpo_concatenado = "\n".join(m["corpo"] for m in mensagens)
        resumo = gerar_resumo(corpo_concatenado)

        # Pessoa-chave em destaque: prioriza um remetente cadastrado.
        with self._lock:
            pessoas_cadastradas = {_normalizar(p) for p in self.pessoas_chave}
        destaque = None
        for m in mensagens:
            if _normalizar(m["remetente"]) in pessoas_cadastradas:
                destaque = m["remetente_nome"] or m["remetente"]
                break
        if destaque is None:
            destaque = ultima["remetente_nome"] or ultima["remetente"]

        return {
            "chave": topico["chave"],
            "assunto": topico["assunto"],
            "pessoa_destaque": destaque,
            "participantes": sorted(topico["pessoas"]),
            "qtd_mensagens": len(mensagens),
            "resumo": resumo,
            "motivos": sorted(topico["motivos"]),
            "ultima_atualizacao": ultima["recebido"].strftime("%d/%m/%Y %H:%M"),
            # Campo auxiliar (ISO) usado só para ordenar no servidor.
            "ultima_atualizacao_ord": ultima["recebido"].isoformat(),
        }
