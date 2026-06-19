# -*- coding: utf-8 -*-
"""
genai_backend.py
================

Backend OPCIONAL de resumo/resposta usando a biblioteca ``google-genai``
(Gemini). Encapsulado e isolado de propósito:

  * **Desligado por padrão.** Só ativa se houver credenciais no ambiente
    (chave de API ou configuração Vertex AI). Sem isso, o app continua
    100% local e nada é enviado para fora — importante para compliance.
  * **À prova de falha.** Qualquer erro (import, rede/proxy, auth, quota)
    faz as funções retornarem ``None``, e o chamador cai no resumo local.

Ativação (defina ANTES de iniciar o app):

  Opção A — Gemini API por chave:
      set GOOGLE_API_KEY=...           (ou GEMINI_API_KEY=...)

  Opção B — Vertex AI (projeto GCP corporativo aprovado):
      set GOOGLE_GENAI_USE_VERTEXAI=1
      set GOOGLE_CLOUD_PROJECT=seu-projeto
      set GOOGLE_CLOUD_LOCATION=us-central1

  Desligar à força (mesmo com credenciais):  set TC_USE_GENAI=0
  Trocar o modelo:                           set TC_GENAI_MODEL=gemini-2.5-flash
"""

from __future__ import annotations

import os
import re
from typing import Any, Dict, List, Optional

try:
    from google import genai  # type: ignore
    _IMPORT_OK = True
    _IMPORT_ERRO = ""
except Exception as _exc:
    genai = None  # type: ignore
    _IMPORT_OK = False
    _IMPORT_ERRO = f"{type(_exc).__name__}: {_exc}"

_MODELO = os.environ.get("TC_GENAI_MODEL", "gemini-2.5-flash")
_cliente: Optional[Any] = None
_cliente_tentado = False
# Após N falhas seguidas, desiste de tentar (evita travar o worker a cada poll
# num ambiente onde o Google está bloqueado pelo proxy).
_falhas = 0
_MAX_FALHAS = 3


def _credenciais_presentes() -> bool:
    if os.environ.get("TC_USE_GENAI", "").lower() in ("0", "false", "no"):
        return False
    if os.environ.get("GOOGLE_GENAI_USE_VERTEXAI", "").lower() in ("1", "true", "yes"):
        return True
    return bool(os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY"))


def ativo() -> bool:
    """True se a biblioteca importou, há credenciais e não estouramos as falhas."""
    return _IMPORT_OK and _credenciais_presentes() and _falhas < _MAX_FALHAS


def _modo() -> str:
    if os.environ.get("GOOGLE_GENAI_USE_VERTEXAI", "").lower() in ("1", "true", "yes"):
        proj = os.environ.get("GOOGLE_CLOUD_PROJECT", "(sem GOOGLE_CLOUD_PROJECT)")
        loc = os.environ.get("GOOGLE_CLOUD_LOCATION", "(sem GOOGLE_CLOUD_LOCATION)")
        return f"Vertex AI (projeto={proj}, location={loc})"
    return "Gemini API (chave)"


def log_estado() -> None:
    """Imprime no console o estado do backend Gemini (chamado no boot do worker)."""
    if not _IMPORT_OK:
        print(f"[genai] biblioteca google-genai NÃO importada -> {_IMPORT_ERRO} "
              "| usando resumo LOCAL.", flush=True)
        return
    if os.environ.get("TC_USE_GENAI", "").lower() in ("0", "false", "no"):
        print("[genai] biblioteca OK, mas DESLIGADA por TC_USE_GENAI=0 "
              "| usando resumo LOCAL.", flush=True)
        return
    if not _credenciais_presentes():
        print("[genai] biblioteca OK, mas SEM credenciais no ambiente "
              "(defina GOOGLE_API_KEY ou GOOGLE_GENAI_USE_VERTEXAI=1 + projeto) "
              "| usando resumo LOCAL.", flush=True)
        return
    print(f"[genai] biblioteca OK e credenciais presentes | modo={_modo()} | "
          f"modelo={_MODELO} | tentará conectar na 1ª varredura.", flush=True)


def _get_cliente() -> Optional[Any]:
    """Cria (uma vez) o cliente Gemini conforme o ambiente. None se falhar."""
    global _cliente, _cliente_tentado
    if _cliente is not None:
        return _cliente
    if _cliente_tentado:
        return _cliente
    _cliente_tentado = True
    try:
        use_vertex = os.environ.get("GOOGLE_GENAI_USE_VERTEXAI", "").lower() in (
            "1", "true", "yes"
        )
        # timeout curto p/ não pendurar a thread do worker se a rede travar.
        http = {"timeout": 30000}  # ms
        if use_vertex:
            _cliente = genai.Client(http_options=http)  # lê projeto/location do env
        else:
            api_key = os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY")
            _cliente = genai.Client(api_key=api_key, http_options=http)
        print(f"[genai] cliente criado com sucesso | modo={_modo()}", flush=True)
    except Exception as exc:
        _cliente = None
        print(f"[genai] FALHA ao criar o cliente -> {type(exc).__name__}: {exc}",
              flush=True)
    return _cliente


def _gerar(prompt: str) -> Optional[str]:
    """Chama o modelo e devolve o texto, ou None em qualquer falha."""
    global _falhas
    cli = _get_cliente()
    if cli is None:
        return None
    try:
        resp = cli.models.generate_content(model=_MODELO, contents=prompt)
        texto = (getattr(resp, "text", "") or "").strip()
        if texto:
            if _falhas:
                print("[genai] conexão restabelecida.", flush=True)
            _falhas = 0  # sucesso reseta o contador
            return texto
        print("[genai] resposta vazia do modelo (sem texto) -> fallback local.",
              flush=True)
        return None
    except Exception as exc:
        _falhas += 1
        print(f"[genai] FALHA na chamada ({_falhas}/{_MAX_FALHAS}) -> "
              f"{type(exc).__name__}: {exc}", flush=True)
        if _falhas >= _MAX_FALHAS:
            print("[genai] desativando tentativas após falhas seguidas "
                  "| seguindo 100% LOCAL.", flush=True)
        return None


def _montar_chain(mensagens: List[Dict[str, Any]], limite_chars: int = 12000) -> str:
    """Monta o texto da conversa (autor + corpo) limitado em tamanho."""
    partes = []
    for m in mensagens:
        autor = (m.get("remetente_nome") or m.get("remetente") or "?")
        corpo = (m.get("corpo") or "").strip()
        partes.append(f"De {autor}:\n{corpo}")
    texto = "\n\n---\n\n".join(partes)
    return texto[-limite_chars:]  # mantém as mensagens mais recentes


def resumir_chain(mensagens: List[Dict[str, Any]]) -> Optional[List[Dict[str, str]]]:
    """Resume a conversa em bullets via Gemini. None se indisponível/erro.

    Retorna no mesmo formato do resumo local: ``[{"autor": "", "ponto": ...}]``.
    """
    if not ativo() or not mensagens:
        return None
    prompt = (
        "Você é assistente de um operador de mesa (Traffic Control) de "
        "derivativos OTC. Resuma a conversa de e-mail abaixo em português, "
        "em bullets curtos e objetivos. Cada bullet deve capturar um ponto "
        "relevante: pedido, decisão, prazo, pendência ou número (trade, valor, "
        "data). Responda APENAS os bullets, um por linha, sem preâmbulo.\n\n"
        + _montar_chain(mensagens)
    )
    texto = _gerar(prompt)
    if not texto:
        return None
    bullets: List[Dict[str, str]] = []
    for linha in texto.splitlines():
        linha = re.sub(r"^\s*[-*••]\s*", "", linha).strip()
        if linha:
            bullets.append({"autor": "", "ponto": linha})
    return bullets or None


def propor_resposta(assunto: str, corpo_chain: str, destinatario: str = "") -> Optional[str]:
    """Gera um rascunho de resposta via Gemini. None se indisponível/erro."""
    if not ativo():
        return None
    prompt = (
        "Você é assistente de um operador de mesa (Traffic Control). Escreva um "
        "rascunho de resposta de e-mail profissional, objetivo e cordial, em "
        "português, pronto para revisão humana. Não invente fatos.\n\n"
        f"Assunto: {assunto}\nPara: {destinatario}\n\n"
        f"Histórico da conversa:\n{corpo_chain[-12000:]}\n\n"
        "Escreva apenas o corpo do e-mail de resposta."
    )
    return _gerar(prompt)
