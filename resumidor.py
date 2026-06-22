# -*- coding: utf-8 -*-
"""
resumidor.py
============

Motor de RESUMO de conversas de e-mail — 100% LOCAL e determinístico.
Nenhuma chamada de rede, nenhuma IA externa, nenhuma credencial, nada sai
da máquina (importante para DLP/compliance num banco).

Por que existe:
    Em vez de mandar o conteúdo dos e-mails para um LLM, este módulo
    *reproduz, com regras explícitas, o raciocínio* que um modelo faria ao
    resumir uma thread:

      1. LIMPA cada mensagem — remove citações ("De:/From:/On ... wrote:"),
         assinaturas e disclaimers jurídicos (o rodapé padrão do banco).
      2. QUEBRA em frases e PONTUA cada frase por relevância: frequência das
         palavras-chave da própria conversa (TF, sem stopwords), com bônus
         para termos de ação, perguntas, datas/valores e a 1ª frase.
      3. SELECIONA os melhores pontos, REMOVE redundância (frases quase
         iguais via similaridade de Jaccard) e os ATRIBUI a quem falou, em
         ordem cronológica.
      4. EXTRAI uma "visão geral" (síntese de 1 linha) e os "pontos de
         atenção": prazos/datas, valores e IDs, perguntas em aberto e pedidos.

    Resultado: um resumo estruturado, legível e ESTÁVEL — o mesmo conjunto de
    e-mails sempre produz exatamente o mesmo resumo.

Saída de ``resumir_conversa``::

    {
        "visao_geral": "Conversa sobre \"...\" — N mensagens entre A, B.",
        "pontos":  [ {"autor": "Maria", "ponto": "..."}, ... ],
        "atencao": [ {"tipo": "Prazo", "texto": "Maria: ... até sexta"}, ... ],
    }
"""

from __future__ import annotations

import re
import unicodedata
from typing import Any, Dict, List, Set, Tuple

# ===========================================================================
# Léxico
# ===========================================================================
# Stopwords PT/EN: ignoradas ao medir a relevância de uma frase.
_STOPWORDS: Set[str] = {
    "a", "o", "as", "os", "um", "uma", "uns", "umas", "de", "do", "da", "dos",
    "das", "em", "no", "na", "nos", "nas", "por", "para", "pra", "com", "sem",
    "e", "ou", "mas", "que", "se", "ao", "aos", "a", "as", "ser", "foi", "sao",
    "este", "esta", "isso", "isto", "esse", "essa", "como", "mais", "menos",
    "ja", "nao", "sim", "the", "of", "to", "and", "in", "on", "for", "is",
    "are", "was", "this", "that", "with", "seu", "sua", "seus", "suas", "meu",
    "minha", "nos", "eu", "ele", "ela", "lhe", "me", "te", "vos", "todo",
    "toda", "todos", "todas", "muito", "pelo", "pela", "ate", "tambem", "esta",
    "estao", "ter", "tem", "favor", "ola", "prezado", "prezada", "caro", "cara",
    "bom", "boa", "dia", "tarde", "noite", "obrigado", "obrigada", "att",
}

# Termos que sinalizam ação/urgência: dão um "boost" à frase no ranking.
_TERMOS_ACAO: Set[str] = {
    "urgente", "prazo", "hoje", "amanha", "confirmar", "confirmacao",
    "aprovar", "aprovacao", "pendente", "vencimento", "deadline", "favor",
    "erro", "bug", "falha", "critico", "imediato", "asap", "prioridade",
    "liquidacao", "settlement", "trade", "registro", "pagamento", "valor",
    "revisar", "validar", "enviar", "responder", "verificar", "solicito",
}

# Verbos/expressões de PEDIDO (busca por substring em texto normalizado).
_PEDIDO: Tuple[str, ...] = (
    "favor", "por favor", "poderia", "pode confirmar", "podem confirmar",
    "confirmar", "confirma ", "aprovar", "aprova ", "enviar", "envie ",
    "revisar", "revise", "precisamos", "preciso ", "solicito", "necessario",
    "verificar", "retornar", "responder", "validar", "providenciar", "gentileza",
)

# ===========================================================================
# Expressões regulares
# ===========================================================================
# CORTE de ruído: tudo a partir do 1º marcador é citação/assinatura/disclaimer.
_RE_CORTE = re.compile(
    r"(?im)^\s*(?:"
    r"de\s*:|from\s*:|enviada?\s+em\s*:|sent\s*:|para\s*:|to\s*:|cc\s*:|"
    r"assunto\s*:|subject\s*:|"
    r"-{3,}|_{3,}|\*{3,}|"
    r"em\s+.{0,80}\s+escreveu\s*:|on\s+.{0,80}\s+wrote\s*:|"
    r"atenciosamente|atensiosamente|att\.?\s*[,:]|abra[cs]os?\b|"
    r"best\s+regards|kind\s+regards|regards\s*[,:]|cumprimentos|saudacoes|"
    r"this\s+(?:e-?mail|message|communication)|esta\s+mensagem|este\s+e-?mail|"
    r"aviso\s+de\s+confidencial|confidentiality\s+notice|disclaimer|"
    r"j\.?p\.?\s*morgan\s+is"
    r")"
)

# DATA/HORA (roda sobre texto NORMALIZADO: sem acento, minúsculo).
_RE_DATA = re.compile(
    r"\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b"
    r"|\b(?:hoje|amanha|ontem|segunda|terca|quarta|quinta|sexta|sabado|domingo)\b"
    r"|\b\d{1,2}h(?:\d{2})?\b"
)
# PRAZO/urgência (texto NORMALIZADO).
_RE_PRAZO = re.compile(
    r"\b(?:prazo|vencimento|deadline|eod|fim do dia|urgente|asap|imediat|"
    r"ate\s|hoje|amanha)\b"
)
# VALORES/percentuais/IDs (roda sobre o texto ORIGINAL, p/ manter R$, %, etc.).
_RE_VALOR = re.compile(
    r"(?:R\$|US\$|USD|BRL|EUR|GBP|CHF)\s?[\d.,]+"
    r"|\b\d[\d.,]*\s?(?:mil|mi|milh[oõ]es|bi|bilh[oõ]es|k|mm|bp|bps)\b"
    r"|\b\d{1,3}(?:[.,]\d+)?\s?%",
    re.IGNORECASE,
)
# Keyword (palavra inteira, com \b nas duas pontas) seguida de um código.
# O \b final evita casar "id" dentro de "Identifiquei", "os" em "ostentar" etc.
_RE_ID = re.compile(
    r"\b(?:trade|deal|book|id|ref|ticket|chamado|nota)\b\s*[:#]?\s*[A-Za-z0-9\-]{3,}\b",
    re.IGNORECASE,
)


# ===========================================================================
# Utilitários
# ===========================================================================
def _norm(texto: str) -> str:
    """Minúsculas, sem acentos (para comparação/regex insensível)."""
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower()


def _limpar(corpo: str) -> str:
    """Remove citações/assinaturas/disclaimers e normaliza espaços."""
    if not corpo:
        return ""
    m = _RE_CORTE.search(corpo)
    if m:
        corpo = corpo[: m.start()]
    return re.sub(r"\s+", " ", corpo).strip()


def _frases(texto: str) -> List[str]:
    """Quebra o texto em frases (por pontuação e quebras de linha)."""
    if not texto:
        return []
    partes = re.split(r"(?<=[.!?])\s+|\n+", texto)
    return [p.strip() for p in partes if len(p.strip()) > 2]


def _tokens(texto: str) -> List[str]:
    """Palavras relevantes (sem stopwords, len>2) do texto normalizado."""
    return [
        t for t in re.findall(r"[a-z0-9à-ú]+", _norm(texto))
        if len(t) > 2 and t not in _STOPWORDS
    ]


def _limitar(texto: str, limite: int = 240) -> str:
    """Corta sem quebrar palavra; remove pontuação solta no fim."""
    texto = texto.strip()
    if len(texto) <= limite:
        return texto
    corte = texto[:limite].rsplit(" ", 1)[0]
    return corte.rstrip(" ,;:.") + "…"


# ===========================================================================
# Seleção de pontos (extrativa, com remoção de redundância)
# ===========================================================================
def _selecionar(
    candidatas: List[Tuple[float, int, int, str, str]], n: int
) -> List[Dict[str, str]]:
    """Pega as ``n`` melhores frases evitando repetir conteúdo (Jaccard).

    ``candidatas``: lista de ``(score, idx_msg, idx_frase, autor, frase)``.
    Retorna ``[{"autor", "ponto"}]`` em ORDEM CRONOLÓGICA (msg, frase).
    """
    ordenadas = sorted(candidatas, key=lambda x: x[0], reverse=True)
    escolhidas: List[Tuple[int, int, str, str, Set[str]]] = []
    for _score, i, j, autor, frase in ordenadas:
        tset = set(_tokens(frase))
        if not tset:
            continue
        redundante = False
        for _, _, _, _, tset2 in escolhidas:
            uniao = len(tset | tset2)
            if uniao and len(tset & tset2) / uniao > 0.55:
                redundante = True
                break
        if redundante:
            continue
        escolhidas.append((i, j, autor, frase, tset))
        if len(escolhidas) >= n:
            break
    escolhidas.sort(key=lambda x: (x[0], x[1]))  # cronológico
    return [{"autor": a, "ponto": _limitar(f, 240)} for _, _, a, f, _ in escolhidas]


def _pontos_atencao(msgs: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """Extrai prazos, valores/IDs, perguntas e pedidos da conversa."""
    itens: List[Dict[str, str]] = []
    vistos: Set[str] = set()
    for m in msgs:
        for fr in _frases(m["corpo"]):
            n = _norm(fr)
            if _RE_PRAZO.search(n) or _RE_DATA.search(n):
                tipo = "Prazo"
            elif any(p in n for p in _PEDIDO):
                tipo = "Pedido"
            elif "?" in fr:
                tipo = "Pergunta"
            elif _RE_VALOR.search(fr) or _RE_ID.search(fr):
                tipo = "Número"
            else:
                continue
            chave = n[:90]
            if chave in vistos:
                continue
            vistos.add(chave)
            prefixo = f"{m['autor']}: " if m["autor"] else ""
            itens.append({"tipo": tipo, "texto": _limitar(prefixo + fr.strip(), 200)})
            if len(itens) >= 6:
                return itens
    return itens


def _visao_geral(
    assunto: str, n_msgs: int, participantes: List[str]
) -> str:
    """Cabeçalho de síntese: assunto, nº de mensagens e quem participou."""
    assunto = (assunto or "").strip() or "(sem assunto)"
    nomes: List[str] = []
    for nm in participantes or []:
        primeiro = nm if "@" in nm else (nm.split(" ")[0] if nm else "")
        if primeiro and primeiro not in nomes:
            nomes.append(primeiro)
    if not nomes:
        quem = ""
    elif len(nomes) == 1:
        quem = f" com {nomes[0]}"
    elif len(nomes) <= 3:
        quem = f" entre {', '.join(nomes)}"
    else:
        quem = f" entre {', '.join(nomes[:3])} +{len(nomes) - 3}"
    plural = "mensagem" if n_msgs == 1 else "mensagens"
    return f'Conversa sobre "{assunto}" — {n_msgs} {plural}{quem}.'


# ===========================================================================
# API pública
# ===========================================================================
def resumir_conversa(
    mensagens: List[Dict[str, Any]],
    assunto: str = "",
    participantes: List[str] | None = None,
) -> Dict[str, Any]:
    """Resume a CHAIN inteira de e-mails em um bloco estruturado e local.

    Args:
        mensagens: lista cronológica; cada item com ``corpo`` e
            ``remetente_nome``/``remetente``.
        assunto: assunto do tópico (para a visão geral).
        participantes: nomes/e-mails dos envolvidos (para a visão geral).

    Returns:
        ``{"visao_geral": str, "pontos": [...], "atencao": [...]}``.
        Nunca lança exceção — em qualquer canto vazio retorna estrutura válida.
    """
    # 1) Limpa cada mensagem e guarda o autor (1º nome).
    msgs: List[Dict[str, str]] = []
    for m in mensagens or []:
        corpo = _limpar(m.get("corpo", ""))
        if not corpo:
            continue
        nome = (m.get("remetente_nome") or m.get("remetente") or "").strip()
        autor = nome.split(" ")[0] if nome and "@" not in nome else nome
        msgs.append({"autor": autor, "corpo": corpo})

    if not msgs:
        return {
            "visao_geral": "Sem conteúdo legível para resumir.",
            "pontos": [],
            "atencao": [],
        }

    # 2) Frequência global das palavras (TF sobre a conversa toda).
    freq: Dict[str, int] = {}
    for m in msgs:
        for tk in _tokens(m["corpo"]):
            freq[tk] = freq.get(tk, 0) + 1

    # 3) Pontua cada frase de cada mensagem.
    candidatas: List[Tuple[float, int, int, str, str]] = []
    for i, m in enumerate(msgs):
        for j, fr in enumerate(_frases(m["corpo"])):
            toks = _tokens(fr)
            if not toks:
                continue
            base = sum(freq.get(t, 0) for t in toks) / (len(toks) ** 0.5)
            n = _norm(fr)
            bonus = 1.0
            if any(t in _TERMOS_ACAO for t in toks):
                bonus *= 1.35
            if "?" in fr:
                bonus *= 1.20
            if _RE_DATA.search(n) or _RE_VALOR.search(fr):
                bonus *= 1.25
            if j == 0:
                bonus *= 1.12  # 1ª frase costuma trazer o assunto da mensagem
            candidatas.append((base * bonus, i, j, m["autor"], fr.strip()))

    # 4) Seleciona os melhores pontos (mais densos quanto maior a thread).
    quantos = min(max(3, len(msgs) + 1), 7)
    pontos = _selecionar(candidatas, quantos)

    # 5) Visão geral + pontos de atenção.
    visao = _visao_geral(assunto, len(msgs), participantes or [])
    atencao = _pontos_atencao(msgs)

    return {"visao_geral": visao, "pontos": pontos, "atencao": atencao}
