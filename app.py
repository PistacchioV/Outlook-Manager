# -*- coding: utf-8 -*-
"""
app.py
======

Servidor Flask do "Gerenciador de E-mails Inteligente".

Expõe:
    * ``/``                      -> painel web (templates/index.html)
    * ``GET  /api/config``       -> configuração atual (palavras/pessoas/intervalo)
    * ``POST /api/palavras``     -> adiciona palavra-chave   {"valor": "..."}
    * ``DELETE /api/palavras``   -> remove palavra-chave      {"valor": "..."}
    * ``POST /api/pessoas``      -> adiciona pessoa-chave     {"valor": "..."}
    * ``DELETE /api/pessoas``    -> remove pessoa-chave        {"valor": "..."}
    * ``POST /api/intervalo``    -> ajusta intervalo (s)      {"segundos": 300}
    * ``GET  /api/topicos``      -> tópicos agrupados (JSON p/ o polling)
    * ``POST /api/varrer-agora`` -> dispara uma varredura imediata

O ``OutlookManager`` é instanciado uma única vez e seu worker é iniciado
junto com a aplicação.
"""

from __future__ import annotations

import os

from flask import Flask, jsonify, render_template, request

from outlook_manager import OutlookManager

app = Flask(__name__)

# Instância única do gerenciador (worker varre a cada 5 min por padrão).
manager = OutlookManager(intervalo_segundos=300)

# Sobe a thread de varredura JÁ NA IMPORTAÇÃO do módulo — assim o worker roda
# independentemente de como o app é iniciado (python app.py, flask run, um
# servidor WSGI como o waitress, ou o "Run" de uma IDE). iniciar_worker() é
# idempotente, então chamar de novo em main() não cria thread duplicada.
manager.iniciar_worker()


# ---------------------------------------------------------------------------
# Página principal
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    """Renderiza o painel web."""
    return render_template("index.html")


# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------
@app.route("/api/config", methods=["GET"])
def get_config():
    return jsonify(manager.get_config())


@app.route("/api/palavras", methods=["POST", "DELETE"])
def palavras():
    valor = (request.get_json(silent=True) or {}).get("valor", "")
    if request.method == "POST":
        manager.add_palavra(valor)
    else:
        manager.remove_palavra(valor)
    return jsonify(manager.get_config())


@app.route("/api/pessoas", methods=["POST", "DELETE"])
def pessoas():
    valor = (request.get_json(silent=True) or {}).get("valor", "")
    if request.method == "POST":
        manager.add_pessoa(valor)
    else:
        manager.remove_pessoa(valor)
    return jsonify(manager.get_config())


@app.route("/api/intervalo", methods=["POST"])
def intervalo():
    segundos = (request.get_json(silent=True) or {}).get("segundos", 300)
    try:
        manager.set_intervalo(int(segundos))
    except (TypeError, ValueError):
        return jsonify({"erro": "intervalo inválido"}), 400
    return jsonify(manager.get_config())


@app.route("/api/conta", methods=["POST"])
def conta():
    """Define a conta/mailbox do Outlook a ser lida (SMTP)."""
    valor = (request.get_json(silent=True) or {}).get("valor", "")
    manager.set_conta(valor)
    return jsonify(manager.get_config())


# ---------------------------------------------------------------------------
# Tópicos / varredura
# ---------------------------------------------------------------------------
@app.route("/api/topicos", methods=["GET"])
def topicos():
    """Endpoint consumido pelo polling do front-end."""
    return jsonify(manager.get_topicos())


@app.route("/api/varrer-agora", methods=["POST"])
def varrer_agora():
    manager.forcar_varredura()
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# Bootstrap da aplicação
# ---------------------------------------------------------------------------
def main() -> None:
    # Inicia a thread de varredura antes de subir o servidor.
    manager.iniciar_worker()
    try:
        # debug=False em ambiente corporativo: o debugger do Werkzeug permite
        # execução remota de código e NÃO deve ficar exposto.
        # use_reloader=False evita que o worker suba duas vezes (o reloader
        # do Flask cria um processo filho que duplicaria a thread COM).
        # host=127.0.0.1 mantém o painel acessível apenas na própria máquina.
        # Porta via env var PORT (padrão 5000). No macOS a 5000 é ocupada pelo
        # AirPlay Receiver (Control Center); rode com PORT=5001 para evitar.
        porta = int(os.environ.get("PORT", "5000"))
        app.run(host="127.0.0.1", port=porta, debug=False, use_reloader=False)
    finally:
        manager.parar_worker()


if __name__ == "__main__":
    main()
