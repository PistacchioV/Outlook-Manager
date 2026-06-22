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
    resumir uma thread. O pipeline é uma versão enxuta de TF-IDF + LexRank
    (centralidade de frases), tudo em Python puro:

      1. LIMPA cada mensagem — remove citações ("De:/From:/On ... wrote:"),
         assinaturas e disclaimers jurídicos (o rodapé padrão do banco).
      2. SEGMENTA em frases (respeitando abreviações, decimais, R$, horas) e
         PONTUA cada frase combinando:
           • informatividade — peso TF-IDF dos termos (palavras raras e
             específicas da conversa pesam mais que as onipresentes);
           • centralidade — quanto a frase "representa" o resto da thread
             (soma das similaridades de cosseno com as demais frases);
           • pistas — bônus para termos de ação, perguntas, datas/valores e
             a 1ª frase de cada mensagem.
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

import math
import re
import unicodedata
from typing import Any, Dict, List, NamedTuple, Sequence, Set, Tuple

# ===========================================================================
# Parâmetros de ajuste (num só lugar, fáceis de calibrar)
# ===========================================================================
_MIN_FRASE = 3          # nº mínimo de caracteres para uma frase valer
_MAX_PONTOS = 7         # teto de pontos no resumo
_MIN_PONTOS = 3         # piso de pontos (quando há conteúdo suficiente)
_MAX_ATENCAO = 6        # teto de "pontos de atenção"
_LIM_PONTO = 240        # corte de caracteres por ponto
_LIM_ATENCAO = 200      # corte de caracteres por ponto de atenção
_JACCARD_PONTO = 0.55   # acima disso, dois pontos são "a mesma coisa"
_JACCARD_ATENCAO = 0.60
_SIM_MIN = 0.12         # similaridade mínima p/ contar na centralidade
_PESO_CENTRAL = 0.55    # peso da centralidade no score final
_PESO_RELEV = 0.45      # peso da informatividade no score final

# ===========================================================================
# Léxico
# ===========================================================================
# Stopwords PT/EN: ignoradas ao medir a relevância de uma frase.
_STOPWORDS: Set[str] = {
    "a", "o", "as", "os", "um", "uma", "uns", "umas", "de", "do", "da", "dos",
    "das", "em", "no", "na", "nos", "nas", "por", "para", "pra", "com", "sem",
    "e", "ou", "mas", "que", "se", "ao", "aos", "ser", "foi", "sao", "este",
    "esta", "isso", "isto", "esse", "essa", "como", "mais", "menos", "ja",
    "nao", "sim", "the", "of", "to", "and", "in", "on", "for", "is", "are",
    "was", "this", "that", "with", "seu", "sua", "seus", "suas", "meu",
    "minha", "eu", "ele", "ela", "lhe", "me", "te", "vos", "todo", "toda",
    "todos", "todas", "muito", "pelo", "pela", "ate", "tambem", "estao",
    "ter", "tem", "ola", "prezado", "prezada", "caro", "cara", "bom", "boa",
    "dia", "tarde", "noite", "obrigado", "obrigada", "att", "sobre", "ainda",
    "quando", "onde", "qual", "quais", "porque", "entao", "assim", "ser",
}

# Termos que sinalizam ação/urgência: dão um "boost" à frase no ranking.
_TERMOS_ACAO: Set[str] = {
    "urgente", "prazo", "hoje", "amanha", "confirmar", "confirmacao",
    "aprovar", "aprovacao", "pendente", "vencimento", "deadline", "favor",
    "erro", "bug", "falha", "critico", "imediato", "asap", "prioridade",
    "liquidacao", "settlement", "trade", "registro", "pagamento", "valor",
    "revisar", "validar", "enviar", "responder", "verificar", "solicito",
    "bloqueio", "bloqueado", "atraso", "atrasado", "reprovado", "rejeitado",
}

# Verbos/expressões de PEDIDO (busca por substring em texto normalizado).
_PEDIDO: Tuple[str, ...] = (
    "por favor", "poderia", "pode confirmar", "podem confirmar", "favor ",
    "confirmar", "confirma ", "aprovar", "aprova ", "enviar", "envie ",
    "revisar", "revise", "precisamos", "preciso ", "solicito", "necessario",
    "verificar", "retornar", "responder", "validar", "providenciar",
    "gentileza", "aguardo", "pode me", "poderiam",
)

# Abreviações comuns: o ponto NÃO encerra a frase depois delas.
_ABREV: Set[str] = {
    "sr", "sra", "srs", "sras", "dr", "dra", "prof", "profa", "ltda", "etc",
    "ex", "obs", "ref", "fig", "pag", "pags", "art", "vs", "aprox", "tel",
    "no", "ph", "mr", "mrs", "ms", "inc", "corp", "dept", "jan", "fev", "mar",
    "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez",
}

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

# Candidato a fim de frase: pontuação terminal seguida de espaço+algo ou fim.
_RE_FIM_FRASE = re.compile(r"[.!?…]+(?=\s+\S|\s*$)")
_RE_TOKEN = re.compile(r"[a-z0-9à-ú]+")

# DATA/HORA (roda sobre texto NORMALIZADO: sem acento, minúsculo).
_RE_DATA = re.compile(
    r"\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b"
    r"|\b(?:hoje|amanha|ontem|segunda|terca|quarta|quinta|sexta|sabado|domingo)"
    r"(?:-feira)?\b"
    r"|\b\d{1,2}h(?:\d{2})?\b"
)
# PRAZO/urgência (texto NORMALIZADO).
_RE_PRAZO = re.compile(
    r"\b(?:prazo|vencimento|deadline|eod|fim do dia|urgente|asap|imediat|"
    r"ate\s|hoje|amanha)\b"
)
# VALORES/percentuais (roda sobre o texto ORIGINAL, p/ manter R$, %, etc.).
_RE_VALOR = re.compile(
    r"(?:R\$|US\$|USD|BRL|EUR|GBP|CHF)\s?[\d.,]+"
    r"|\b\d[\d.,]*\s?(?:mil|mi|milh[oõ]es|bi|bilh[oõ]es|k|mm|bp|bps)\b"
    r"|\b\d{1,3}(?:[.,]\d+)?\s?%",
    re.IGNORECASE,
)
# Keyword (palavra inteira) seguida de um código — ex.: "trade 12345", "Ref: AB9".
# O \b final evita casar "id" dentro de "Identifiquei", etc.
_RE_ID = re.compile(
    r"\b(?:trade|deal|book|id|ref|ticket|chamado|nota)\b\s*[:#]?\s*[A-Za-z0-9\-]{3,}\b",
    re.IGNORECASE,
)


# ===========================================================================
# Estrutura interna
# ===========================================================================
class _Frase(NamedTuple):
    """Uma frase já segmentada e pré-processada (tudo computado uma vez só)."""
    idx_msg: int          # ordem da mensagem na thread
    idx_frase: int        # ordem da frase dentro da mensagem
    autor: str
    texto: str            # original (trimmed) — vai para a saída
    norm: str             # normalizado — para regex/comparação
    tokens: Tuple[str, ...]

    @property
    def token_set(self) -> Set[str]:
        return set(self.tokens)


# ===========================================================================
# Utilitários de texto
# ===========================================================================
def _norm(texto: str) -> str:
    """Minúsculas, sem acentos (para comparação/regex insensível)."""
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower()


def _primeiro_nome(nome: str) -> str:
    """1º nome de quem assina; e-mails ficam inteiros (não têm 1º nome)."""
    nome = (nome or "").strip()
    if not nome or "@" in nome:
        return nome
    return nome.split()[0]


def _limpar(corpo: str) -> str:
    """Remove citações/assinaturas/disclaimers e normaliza espaços."""
    if not corpo:
        return ""
    m = _RE_CORTE.search(corpo)
    if m:
        corpo = corpo[: m.start()]
    return re.sub(r"[ \t]+", " ", corpo).strip()


def _segmentar(texto: str) -> List[str]:
    """Quebra o texto em frases respeitando abreviações, decimais e horas.

    Estratégia: quebra de linha sempre separa; dentro de cada linha, só corta
    em ``.!?…`` quando há espaço depois (logo "R$ 1.500" e "3.5%" não cortam)
    e quando a palavra anterior não é abreviação nem uma inicial isolada.
    """
    if not texto:
        return []
    frases: List[str] = []
    for linha in re.split(r"[\n\r]+", texto):
        inicio = 0
        for m in _RE_FIM_FRASE.finditer(linha):
            anterior = linha[inicio:m.start()].rstrip()
            ultima = re.split(r"[\s(]+", anterior)[-1] if anterior else ""
            base = _norm(ultima.strip(".")).strip()
            if base in _ABREV or (len(base) <= 1 and base.isalpha()):
                continue
            frases.append(linha[inicio:m.end()].strip())
            inicio = m.end()
        resto = linha[inicio:].strip()
        if resto:
            frases.append(resto)
    return [f for f in frases if len(f) > _MIN_FRASE]


def _tokens(norm_texto: str) -> List[str]:
    """Palavras relevantes (sem stopwords, len>2) de um texto JÁ normalizado."""
    return [t for t in _RE_TOKEN.findall(norm_texto)
            if len(t) > 2 and t not in _STOPWORDS]


def _limitar(texto: str, limite: int) -> str:
    """Corta sem quebrar palavra; remove pontuação solta no fim."""
    texto = texto.strip()
    if len(texto) <= limite:
        return texto
    corte = texto[:limite].rsplit(" ", 1)[0]
    return corte.rstrip(" ,;:.") + "…"


def _jaccard(a: Set[str], b: Set[str]) -> float:
    """Similaridade de Jaccard entre dois conjuntos de tokens."""
    if not a or not b:
        return 0.0
    uniao = len(a | b)
    return len(a & b) / uniao if uniao else 0.0


# ===========================================================================
# Pré-processamento da thread
# ===========================================================================
def _preparar_frases(mensagens: Sequence[Dict[str, Any]]) -> List[_Frase]:
    """Limpa as mensagens e devolve a lista plana de frases pré-processadas."""
    frases: List[_Frase] = []
    idx_msg = 0
    for m in mensagens or []:
        corpo = _limpar(str(m.get("corpo", "") or ""))
        if not corpo:
            continue
        autor = _primeiro_nome(m.get("remetente_nome") or m.get("remetente") or "")
        for j, fr in enumerate(_segmentar(corpo)):
            norm = _norm(fr)
            toks = tuple(_tokens(norm))
            if not toks:
                continue
            frases.append(_Frase(idx_msg, j, autor, fr.strip(), norm, toks))
        idx_msg += 1
    return frases


def _idf(frases: Sequence[_Frase]) -> Dict[str, float]:
    """IDF com a MENSAGEM como documento: termo onipresente pesa menos."""
    docs: Dict[int, Set[str]] = {}
    for f in frases:
        docs.setdefault(f.idx_msg, set()).update(f.tokens)
    n_docs = max(1, len(docs))
    df: Dict[str, int] = {}
    for tokens in docs.values():
        for t in tokens:
            df[t] = df.get(t, 0) + 1
    return {t: math.log(1 + n_docs / (1 + d)) for t, d in df.items()}


def _vetor(tokens: Sequence[str], idf: Dict[str, float]) -> Tuple[Dict[str, float], float]:
    """Vetor TF-IDF (esparso) de uma frase + sua magnitude (para cosseno)."""
    vec: Dict[str, float] = {}
    for t in tokens:
        vec[t] = vec.get(t, 0.0) + idf.get(t, 0.0)
    mag = math.sqrt(sum(w * w for w in vec.values()))
    return vec, mag


def _cosseno(va: Dict[str, float], na: float,
             vb: Dict[str, float], nb: float) -> float:
    """Similaridade de cosseno entre dois vetores esparsos."""
    if na == 0.0 or nb == 0.0:
        return 0.0
    if len(va) > len(vb):  # itera sobre o menor
        va, vb = vb, va
    dot = sum(w * vb.get(t, 0.0) for t, w in va.items())
    return dot / (na * nb)


def _minmax(valores: List[float]) -> List[float]:
    """Normaliza para [0,1]; tudo igual → 0.5 (deixa o bônus decidir)."""
    if not valores:
        return []
    lo, hi = min(valores), max(valores)
    span = hi - lo
    if span <= 0:
        return [0.5] * len(valores)
    return [(v - lo) / span for v in valores]


def _bonus_pistas(f: _Frase) -> float:
    """Multiplicador por pistas de relevância (ação, pergunta, data/valor)."""
    bonus = 1.0
    if any(t in _TERMOS_ACAO for t in f.tokens):
        bonus *= 1.35
    if "?" in f.texto:
        bonus *= 1.20
    if _RE_DATA.search(f.norm) or _RE_VALOR.search(f.texto):
        bonus *= 1.25
    if f.idx_frase == 0:
        bonus *= 1.12  # 1ª frase costuma trazer o assunto da mensagem
    return bonus


# ===========================================================================
# Ranqueamento e seleção
# ===========================================================================
def _pontuar(frases: List[_Frase], idf: Dict[str, float]) -> List[float]:
    """Score por frase = (centralidade + informatividade) × pistas."""
    n = len(frases)
    vetores = [_vetor(f.tokens, idf) for f in frases]

    # Centralidade (LexRank-lite): soma das similaridades acima do limiar.
    central = [0.0] * n
    for a in range(n):
        va, na = vetores[a]
        for b in range(a + 1, n):
            vb, nb = vetores[b]
            sim = _cosseno(va, na, vb, nb)
            if sim >= _SIM_MIN:
                central[a] += sim
                central[b] += sim
    if n > 1:
        central = [c / (n - 1) for c in central]

    # Informatividade: peso IDF dos termos, normalizado pelo tamanho da frase.
    relev = [sum(idf.get(t, 0.0) for t in f.tokens) / (len(f.tokens) ** 0.5)
             for f in frases]

    nc, nr = _minmax(central), _minmax(relev)
    return [(_PESO_CENTRAL * nc[i] + _PESO_RELEV * nr[i]) * _bonus_pistas(f)
            for i, f in enumerate(frases)]


def _selecionar(frases: List[_Frase], scores: List[float],
                quantos: int) -> List[Dict[str, str]]:
    """Top-``quantos`` frases sem redundância (Jaccard), em ordem cronológica."""
    ordem = sorted(range(len(frases)),
                   key=lambda i: (-scores[i], frases[i].idx_msg, frases[i].idx_frase))
    escolhidas: List[_Frase] = []
    sets: List[Set[str]] = []
    for i in ordem:
        f = frases[i]
        ts = f.token_set
        if any(_jaccard(ts, s) > _JACCARD_PONTO for s in sets):
            continue
        escolhidas.append(f)
        sets.append(ts)
        if len(escolhidas) >= quantos:
            break
    escolhidas.sort(key=lambda f: (f.idx_msg, f.idx_frase))
    return [{"autor": f.autor, "ponto": _limitar(f.texto, _LIM_PONTO)}
            for f in escolhidas]


def _classificar(f: _Frase) -> str | None:
    """Rotula uma frase como Prazo/Pedido/Pergunta/Número (ou None)."""
    if _RE_PRAZO.search(f.norm) or _RE_DATA.search(f.norm):
        return "Prazo"
    if any(p in f.norm for p in _PEDIDO):
        return "Pedido"
    if "?" in f.texto:
        return "Pergunta"
    if _RE_VALOR.search(f.texto) or _RE_ID.search(f.texto):
        return "Número"
    return None


def _pontos_atencao(frases: List[_Frase]) -> List[Dict[str, str]]:
    """Extrai prazos, valores/IDs, perguntas e pedidos — sem repetir conteúdo."""
    itens: List[Dict[str, str]] = []
    sets: List[Set[str]] = []
    for f in frases:  # já em ordem cronológica
        tipo = _classificar(f)
        if tipo is None:
            continue
        ts = f.token_set
        if any(_jaccard(ts, s) > _JACCARD_ATENCAO for s in sets):
            continue
        sets.append(ts)
        prefixo = f"{f.autor}: " if f.autor else ""
        itens.append({"tipo": tipo,
                      "texto": _limitar(prefixo + f.texto, _LIM_ATENCAO)})
        if len(itens) >= _MAX_ATENCAO:
            break
    return itens


def _visao_geral(assunto: str, n_msgs: int, participantes: Sequence[str]) -> str:
    """Cabeçalho de síntese: assunto, nº de mensagens e quem participou."""
    assunto = (assunto or "").strip() or "(sem assunto)"
    nomes: List[str] = []
    for nm in participantes or []:
        primeiro = _primeiro_nome(nm)
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
    frases = _preparar_frases(mensagens)
    n_msgs = (frases[-1].idx_msg + 1) if frases else 0

    if not frases:
        return {
            "visao_geral": _visao_geral(assunto, 0, participantes or []),
            "pontos": [],
            "atencao": [],
        }

    idf = _idf(frases)
    scores = _pontuar(frases, idf)

    quantos = max(_MIN_PONTOS, min(_MAX_PONTOS, n_msgs + 1))
    quantos = min(quantos, len(frases))

    return {
        "visao_geral": _visao_geral(assunto, n_msgs, participantes or []),
        "pontos": _selecionar(frases, scores, quantos),
        "atencao": _pontos_atencao(frases),
    }
