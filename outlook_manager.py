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
    * Geração de um resumo curto por tópico (sumarização extrativa local).

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

# Arquivos de persistência (mesmo diretório do módulo).
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_CONFIG_PATH = os.path.join(_BASE_DIR, "config.json")
# Histórico acumulado de cada tópico (mensagens únicas + trail de resumos).
# Contém corpo de e-mails — NÃO versionar (está no .gitignore).
_HISTORICO_PATH = os.path.join(_BASE_DIR, "historico_topicos.json")

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
# RESUMO AUTOMÁTICO (extrativo, 100% local — sem LLM/chamadas externas)
# ===========================================================================
# Stopwords PT-BR mais comuns: ignoradas ao pontuar a relevância das frases.
_STOPWORDS_PT = {
    "a", "o", "as", "os", "um", "uma", "uns", "umas", "de", "do", "da", "dos",
    "das", "em", "no", "na", "nos", "nas", "por", "para", "pra", "com", "sem",
    "e", "ou", "mas", "que", "se", "ao", "aos", "à", "às", "é", "ser", "foi",
    "são", "este", "esta", "isso", "isto", "esse", "essa", "como", "mais",
    "menos", "já", "não", "sim", "the", "of", "to", "and", "in", "on", "for",
    "seu", "sua", "seus", "suas", "meu", "minha", "nós", "eu", "ele", "ela",
    "lhe", "me", "te", "vos", "todo", "toda", "todos", "todas", "muito",
    "pelo", "pela", "até", "também", "está", "estão", "ter", "tem", "favor",
}

# Termos que sinalizam ação/urgência: dão um leve "boost" à frase no ranking.
_TERMOS_ACAO = {
    "urgente", "prazo", "hoje", "amanha", "amanhã", "confirmar", "confirmação",
    "aprovar", "aprovação", "pendente", "vencimento", "deadline", "favor",
    "erro", "bug", "falha", "crítico", "critico", "imediato", "asap",
    "liquidação", "liquidacao", "settlement", "trade", "registro",
}


def gerar_resumo(texto: str, max_frases: int = 3) -> str:
    """Resume um corpo de e-mail por **sumarização extrativa** (sem LLM).

    Estratégia (clássica, determinística e 100% local):
      1. Limpa o texto (remove cadeias de resposta antigas e assinaturas).
      2. Quebra em frases.
      3. Pontua cada frase pela frequência de suas palavras relevantes
         (ignorando stopwords), normalizada pelo tamanho da frase, com um
         pequeno bônus para a 1ª frase e para frases com termos de ação.
      4. Escolhe as ``max_frases`` melhores e as **reordena pela posição
         original**, para o resumo ler de forma natural.

    Não faz nenhuma chamada externa — adequado a ambientes corporativos
    onde LLMs/APIs externas não são permitidas. Nunca lança exceção.

    Args:
        texto: Corpo do e-mail (texto puro).
        max_frases: Número máximo de frases no resumo.

    Returns:
        Uma string com o resumo.
    """
    if not texto or not texto.strip():
        return "Sem conteúdo para resumir."

    # Normaliza espaços/quebras e corta respostas antigas / assinaturas.
    limpo = re.sub(r"\s+", " ", texto).strip()
    limpo = re.split(r"(?:^|\s)(?:De:|From:|Enviada em:|Sent:|-{3,}|_{3,})", limpo)[0].strip()

    frases = [f.strip() for f in re.split(r"(?<=[.!?])\s+", limpo) if f.strip()]
    if not frases:
        return limpo[:300] + ("..." if len(limpo) > 300 else "")
    if len(frases) <= max_frases:
        return _limitar(limpo)

    # Frequência das palavras relevantes (sobre o texto todo).
    freq: Dict[str, int] = {}
    for palavra in re.findall(r"[a-zà-ú0-9]+", _normalizar(limpo)):
        if len(palavra) <= 2 or palavra in _STOPWORDS_PT:
            continue
        freq[palavra] = freq.get(palavra, 0) + 1

    # Pontua cada frase.
    pontuadas = []
    for idx, frase in enumerate(frases):
        palavras = [
            p for p in re.findall(r"[a-zà-ú0-9]+", _normalizar(frase))
            if len(p) > 2 and p not in _STOPWORDS_PT
        ]
        if not palavras:
            score = 0.0
        else:
            # Soma das frequências, normalizada para não privilegiar frases longas.
            base = sum(freq.get(p, 0) for p in palavras) / (len(palavras) ** 0.5)
            bonus_acao = 1.25 if any(p in _TERMOS_ACAO for p in palavras) else 1.0
            bonus_inicio = 1.15 if idx == 0 else 1.0  # 1ª frase costuma dar o tema
            score = base * bonus_acao * bonus_inicio
        pontuadas.append((score, idx, frase))

    # Top-N por score, depois reordena pela posição original (leitura natural).
    melhores = sorted(pontuadas, key=lambda x: x[0], reverse=True)[:max_frases]
    melhores.sort(key=lambda x: x[1])
    return _limitar(" ".join(f for _, _, f in melhores))


def _limitar(texto: str, limite: int = 320) -> str:
    """Corta o texto em ``limite`` caracteres sem quebrar no meio da palavra."""
    if len(texto) <= limite:
        return texto
    corte = texto[:limite].rsplit(" ", 1)[0]
    return corte.rstrip() + "..."


def propor_resposta(assunto: str, historico_texto: str, destinatario: str = "") -> str:
    """Propõe um rascunho de resposta com base no HISTÓRICO COMPLETO do tópico.

    Implementação **determinística e 100% local** (sem LLM/chamadas externas,
    adequada a ambientes corporativos com restrição de API): detecta sinais no
    histórico — urgência, prazos, perguntas e pedidos de confirmação — e monta
    um rascunho profissional curto em português, pronto para revisão humana.

    Para ajustar o tom/conteúdo, edite as frases de ``linhas`` abaixo e os
    conjuntos de gatilhos (``urgente``/``tem_prazo``/etc.).

    Args:
        assunto: Assunto do tópico.
        historico_texto: Concatenação de TODAS as mensagens do tópico.
        destinatario: Nome da pessoa-chave (saudação personalizada).

    Returns:
        Rascunho de resposta. Nunca lança exceção.
    """
    texto = _normalizar(historico_texto)
    primeiro_nome = (destinatario or "").strip().split(" ")[0] or "tudo bem"
    saudacao = f"Olá {primeiro_nome}," if destinatario else "Olá,"

    # Sinais detectados no histórico.
    urgente = any(p in texto for p in ("urgente", "asap", "imediato", "hoje ainda", "o quanto antes"))
    tem_prazo = any(p in texto for p in ("prazo", "ate sexta", "ate amanha", "vencimento", "deadline", "ate o fim"))
    tem_pergunta = "?" in historico_texto or any(
        p in texto for p in ("voce pode", "poderia", "consegue", "pode confirmar", "qual ", "quando ")
    )
    pede_confirmacao = any(p in texto for p in ("confirmar", "confirma", "de acordo", "aprovar", "aprovacao"))

    linhas = [saudacao, ""]
    linhas.append(f"Obrigado pelo retorno sobre \"{assunto.strip()}\".")

    if urgente:
        linhas.append("Entendi a urgência e já estou priorizando este ponto.")
    if tem_pergunta:
        linhas.append("Sobre os pontos levantados, vou verificar internamente e retorno com os detalhes.")
    if pede_confirmacao:
        linhas.append("Assim que validar, envio a confirmação formal.")
    if tem_prazo:
        linhas.append("Estou atento ao prazo mencionado e darei retorno dentro dele.")
    if not (urgente or tem_pergunta or pede_confirmacao or tem_prazo):
        linhas.append("Fico à disposição para seguir com os próximos passos.")

    linhas.append("")
    linhas.append("Atenciosamente,")
    return "\n".join(linhas)


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
        # Janela de varredura: só considera e-mails dos últimos N dias.
        self.dias_janela: int = 7

        # ----- Estado de resultados -----
        # Dicionário: chave_topico -> dados do tópico (ver _montar_topico).
        self._topicos: Dict[str, Dict[str, Any]] = {}
        self.ultima_varredura: Optional[str] = None
        self.status_worker: str = "parado"
        # Conta que o Outlook REALMENTE reconheceu na última varredura
        # bem-sucedida (None enquanto não conectou / em erro).
        self.conta_conectada: Optional[str] = None

        # Histórico acumulado por tópico: chave -> {assunto, mensagens, resumos}.
        # Persiste em disco e é manipulado apenas na thread do worker.
        self._historico: Dict[str, Dict[str, Any]] = {}
        # Instante-base fixo para o modo simulado (timestamps estáveis entre
        # varreduras, para o dedup do histórico funcionar como no Outlook real).
        self._sim_base = datetime.now()

        # ----- Concorrência -----
        self._lock = threading.Lock()
        self._stop_event = threading.Event()       # sinaliza parada do worker
        self._wake_event = threading.Event()       # sinaliza "varrer agora"
        self._thread: Optional[threading.Thread] = None

        # ----- Persistência -----
        # Carrega config.json e o histórico (se existirem).
        self._carregar_config()
        self._carregar_historico()

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

    def _carregar_historico(self) -> None:
        """Carrega o histórico de tópicos do disco (se existir)."""
        try:
            with open(_HISTORICO_PATH, "r", encoding="utf-8") as fp:
                self._historico = json.load(fp)
        except FileNotFoundError:
            self._historico = {}
        except Exception:
            self._historico = {}  # histórico corrompido não impede o boot

    def _salvar_historico(self) -> None:
        """Grava o histórico em disco (escrita atômica)."""
        try:
            tmp = _HISTORICO_PATH + ".tmp"
            with open(tmp, "w", encoding="utf-8") as fp:
                json.dump(self._historico, fp, ensure_ascii=False, indent=2)
            os.replace(tmp, _HISTORICO_PATH)
        except Exception:
            pass

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
        modo = "SIMULADO" if MODO_SIMULADO else "OUTLOOK (pywin32)"
        print(
            f"[worker] iniciado | modo={modo} | conta-alvo={self.conta_email or '(padrão)'}"
            f" | intervalo={self.intervalo_segundos}s",
            flush=True,
        )

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
        self.status_worker = "iniciando"
        com_iniciado = False

        # Inicializa o COM desta thread (obrigatório no Windows). Falha aqui
        # precisa ficar VISÍVEL no painel — não pode virar um "parado" mudo.
        if not MODO_SIMULADO:
            try:
                pythoncom.CoInitialize()
                com_iniciado = True
            except Exception as exc:  # noqa: BLE001
                self.status_worker = f"erro: falha ao iniciar COM (CoInitialize): {exc}"
                return

        try:
            self.status_worker = "rodando"

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
        except Exception as exc:  # noqa: BLE001  (thread não pode morrer calada)
            self.status_worker = f"erro: worker interrompido: {exc}"
        finally:
            # Só marca "parado" se a parada foi solicitada; caso contrário,
            # preserva a mensagem de erro para aparecer no painel.
            if self._stop_event.is_set():
                self.status_worker = "parado"
            if com_iniciado:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

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
            # Log no console: confirma que a varredura rodou e o que achou.
            print(
                f"[worker] varredura {self.ultima_varredura} | conta="
                f"{self.conta_conectada} | {len(emails)} e-mails lidos | "
                f"{len(self._topicos)} tópico(s) relevante(s)",
                flush=True,
            )
        except Exception as exc:  # noqa: BLE001  (logamos e seguimos)
            self.status_worker = f"erro: {exc}"
            self.conta_conectada = None  # não confirmamos a caixa
            print(f"[worker] ERRO na varredura: {exc}", flush=True)

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
        try:
            itens.Sort("[ReceivedTime]", True)  # mais novos primeiro
        except Exception:
            pass

        # Janela de tempo SEM Restrict por string de data: a string do
        # Restrict é interpretada no FORMATO REGIONAL do Windows (em pt-BR é
        # dd/MM/yyyy), então uma string em formato americano falha ou zera o
        # filtro. Como os itens já vêm ordenados do mais novo para o mais
        # antigo, basta iterar e PARAR ao passar do limite — 100% independente
        # de locale.
        limite = datetime.now() - timedelta(days=self.dias_janela)

        try:
            total_inbox = itens.Count
        except Exception:
            total_inbox = -1
        mais_recente: Optional[datetime] = None

        emails: List[Dict[str, Any]] = []
        vistos = 0
        for item in itens:
            vistos += 1
            if vistos > 3000:          # trava de segurança p/ caixas enormes
                break
            try:
                # 43 == olMail (Class do MailItem). Ignora convites, etc.
                if getattr(item, "Class", 43) != 43:
                    continue

                recebido = self._to_datetime(getattr(item, "ReceivedTime", None))
                if mais_recente is None:
                    mais_recente = recebido
                if recebido < limite:
                    break              # ordenado desc: o resto é mais antigo

                remetente = self._extrair_email_remetente(item)
                emails.append({
                    "assunto": getattr(item, "Subject", "") or "(sem assunto)",
                    "remetente": remetente,
                    "remetente_nome": getattr(item, "SenderName", "") or remetente,
                    "corpo": getattr(item, "Body", "") or "",
                    "recebido": recebido,
                    "conversation_id": getattr(item, "ConversationID", "") or "",
                })
            except Exception:
                continue  # item problemático não derruba a varredura

        recente_str = mais_recente.strftime("%d/%m/%Y %H:%M") if mais_recente else "—"
        print(
            f"[worker] caixa='{self.conta_conectada}' | itens na Inbox={total_inbox}"
            f" | e-mail mais recente={recente_str} | {len(emails)} na janela de "
            f"{self.dias_janela} dias",
            flush=True,
        )
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
        agora = self._sim_base  # base fixa: timestamps estáveis entre varreduras
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

        # Monta a estrutura final serializável (resumo + resposta + metadados).
        finais: Dict[str, Dict[str, Any]] = {}
        for chave, topico in novos_topicos.items():
            finais[chave] = self._montar_topico(topico)

        # Persiste o histórico atualizado (resumos acumulados + mensagens).
        self._salvar_historico()

        with self._lock:
            self._topicos = finais

    @staticmethod
    def _assinatura_msg(email: Dict[str, Any]) -> str:
        """Identidade estável de uma mensagem, para dedup no histórico."""
        recebido = email["recebido"]
        iso = recebido.isoformat() if hasattr(recebido, "isoformat") else str(recebido)
        return f"{_normalizar(email['remetente'])}|{iso}|{len(email.get('corpo',''))}"

    def _atualizar_historico(self, topico: Dict[str, Any]) -> Dict[str, Any]:
        """Funde as mensagens do tópico no histórico persistente e o devolve.

        - Acrescenta apenas mensagens inéditas (dedup por assinatura).
        - Acrescenta um snapshot de resumo (com data) quando ele MUDA.
        Roda apenas na thread do worker (sem necessidade de lock no histórico).
        """
        chave = topico["chave"]
        registro = self._historico.get(chave)
        if registro is None:
            registro = {"assunto": topico["assunto"], "mensagens": [], "resumos": []}
            self._historico[chave] = registro

        registro["assunto"] = topico["assunto"]  # mantém o assunto mais recente
        vistas = {m["assinatura"] for m in registro["mensagens"]}

        for email in sorted(topico["mensagens"], key=lambda m: m["recebido"]):
            sig = self._assinatura_msg(email)
            if sig in vistas:
                continue
            vistas.add(sig)
            registro["mensagens"].append({
                "assinatura": sig,
                "remetente": email["remetente"],
                "remetente_nome": email["remetente_nome"],
                "recebido": email["recebido"].isoformat(),
                "corpo": email["corpo"],
            })

        # Ordena cronologicamente e gera o resumo do HISTÓRICO COMPLETO.
        registro["mensagens"].sort(key=lambda m: m["recebido"])
        corpo_completo = "\n\n".join(m["corpo"] for m in registro["mensagens"])
        resumo_atual = gerar_resumo(corpo_completo)

        # Snapshot de resumo só quando muda (evita inflar o trail a cada poll).
        if not registro["resumos"] or registro["resumos"][-1]["resumo"] != resumo_atual:
            registro["resumos"].append({
                "data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "resumo": resumo_atual,
            })

        return registro

    def _montar_topico(self, topico: Dict[str, Any]) -> Dict[str, Any]:
        """Constrói o dict final de um tópico (pronto para virar JSON).

        Usa o HISTÓRICO COMPLETO acumulado (não só a varredura atual) para o
        resumo e para a resposta sugerida.
        """
        mensagens = sorted(topico["mensagens"], key=lambda m: m["recebido"])
        ultima = mensagens[-1]

        # Funde no histórico e obtém resumo acumulado + trail.
        registro = self._atualizar_historico(topico)
        resumo = registro["resumos"][-1]["resumo"]
        corpo_completo = "\n\n".join(m["corpo"] for m in registro["mensagens"])

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

        # Resposta sugerida com base no histórico completo (heurística local).
        resposta_sugerida = propor_resposta(topico["assunto"], corpo_completo, destaque)

        return {
            "chave": topico["chave"],
            "assunto": topico["assunto"],
            "pessoa_destaque": destaque,
            "participantes": sorted(topico["pessoas"]),
            "qtd_mensagens": len(registro["mensagens"]),  # total acumulado
            "resumo": resumo,
            "resposta_sugerida": resposta_sugerida,
            "historico_resumos": registro["resumos"][-5:],  # últimos snapshots
            "qtd_resumos": len(registro["resumos"]),
            "motivos": sorted(topico["motivos"]),
            "ultima_atualizacao": ultima["recebido"].strftime("%d/%m/%Y %H:%M"),
            # Campo auxiliar (ISO) usado só para ordenar no servidor.
            "ultima_atualizacao_ord": ultima["recebido"].isoformat(),
        }
