# -*- coding: utf-8 -*-
"""
diagnostico.py
==============

Ferramenta de diagnóstico da conexão com o Outlook (Windows).

Rode na MESMA máquina e MESMO Python em que você roda o app:

    python diagnostico.py

Ele testa, em ordem, cada etapa que o app precisa e aponta exatamente onde
falha — com uma sugestão de correção. Não altera nada; só lê.
"""

from __future__ import annotations

import platform
import struct
import sys


def linha(titulo: str) -> None:
    print("\n" + "=" * 64)
    print(titulo)
    print("=" * 64)


def ok(msg: str) -> None:
    print(f"  [OK]    {msg}")


def falha(msg: str) -> None:
    print(f"  [FALHA] {msg}")


def info(msg: str) -> None:
    print(f"          {msg}")


def main() -> None:
    # ---------------------------------------------------------------- #
    # 0) Ambiente Python
    # ---------------------------------------------------------------- #
    linha("0) Ambiente Python")
    print(f"  Python   : {sys.version.split()[0]}")
    print(f"  Executável: {sys.executable}")
    print(f"  Bits     : {struct.calcsize('P') * 8}-bit")
    print(f"  SO       : {platform.system()} {platform.release()}")

    if platform.system() != "Windows":
        falha("Você NÃO está no Windows. O pywin32/Outlook COM só funciona no "
              "Windows com o Outlook Clássico. Rode este script na sua "
              "estação corporativa.")
        return

    # ---------------------------------------------------------------- #
    # 1) pywin32 instalado?
    # ---------------------------------------------------------------- #
    linha("1) Biblioteca pywin32")
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
        ok("pywin32 importado com sucesso.")
    except Exception as exc:
        falha(f"Não foi possível importar o pywin32: {exc}")
        info("CORREÇÃO: instale no MESMO Python do app:")
        info(f"    \"{sys.executable}\" -m pip install pywin32")
        info("Se a empresa bloqueia o PyPI, peça o wheel do pywin32 ao TI.")
        info("OBS: se o app mostra 'Modo simulado' no Windows, é exatamente")
        info("     isto: o pywin32 não está sendo importado neste Python.")
        return

    import pythoncom
    import win32com.client

    # ---------------------------------------------------------------- #
    # 2) Inicializa o COM
    # ---------------------------------------------------------------- #
    linha("2) Inicialização do COM")
    try:
        pythoncom.CoInitialize()
        ok("CoInitialize() funcionou.")
    except Exception as exc:
        falha(f"CoInitialize falhou: {exc}")
        return

    # ---------------------------------------------------------------- #
    # 3) Dispatch do Outlook
    # ---------------------------------------------------------------- #
    linha("3) Conexão com o Outlook (COM Dispatch)")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ok("Outlook.Application conectado.")
    except Exception as exc:
        falha(f"Não foi possível conectar ao Outlook: {exc}")
        info("CAUSAS MAIS COMUNS (ambiente corporativo):")
        info("  • 'Novo Outlook' ligado — ele NÃO suporta COM. Desligue o")
        info("    botão 'Novo Outlook' no canto superior direito do Outlook")
        info("    (volta para o Clássico). O Clássico precisa estar instalado.")
        info("  • Outlook não instalado / não configurado.")
        info("  • Antivírus/EDR ou GPO bloqueando automação COM (fale com o TI).")
        info("  • Python rodando como ADMINISTRADOR e Outlook como usuário")
        info("    normal (ou vice-versa): níveis de elevação diferentes impedem")
        info("    o COM de conectar. Rode os dois no MESMO nível.")
        pythoncom.CoUninitialize()
        return

    # ---------------------------------------------------------------- #
    # 4) Namespace MAPI
    # ---------------------------------------------------------------- #
    linha("4) Sessão MAPI")
    try:
        namespace = outlook.GetNamespace("MAPI")
        ok("GetNamespace('MAPI') funcionou.")
    except Exception as exc:
        falha(f"GetNamespace falhou: {exc}")
        pythoncom.CoUninitialize()
        return

    # ---------------------------------------------------------------- #
    # 5) Contas e Stores disponíveis
    # ---------------------------------------------------------------- #
    linha("5) Contas e caixas (Stores) no perfil")
    contas = []
    try:
        for acc in namespace.Session.Accounts:
            smtp = getattr(acc, "SmtpAddress", "") or "(sem SMTP)"
            contas.append(smtp)
            print(f"  • Conta : {acc.DisplayName}  |  SMTP: {smtp}")
        if not contas:
            falha("Nenhuma conta encontrada no perfil do Outlook.")
    except Exception as exc:
        falha(f"Não foi possível listar contas: {exc}")

    try:
        print("  --- Stores (caixas) ---")
        for st in namespace.Stores:
            print(f"  • Store : {getattr(st, 'DisplayName', '(sem nome)')}")
    except Exception as exc:
        info(f"(não foi possível listar Stores: {exc})")

    # ---------------------------------------------------------------- #
    # 6) Inbox da conta padrão
    # ---------------------------------------------------------------- #
    linha("6) Inbox da conta padrão (GetDefaultFolder)")
    try:
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        total = inbox.Items.Count
        ok(f"Inbox padrão acessível — {total} itens.")
        try:
            ok(f"Store da conta padrão: {inbox.Store.DisplayName}")
        except Exception:
            pass
    except Exception as exc:
        falha(f"Não foi possível abrir a Inbox padrão: {exc}")

    # ---------------------------------------------------------------- #
    # 7) Conta configurada no app (config.json)
    # ---------------------------------------------------------------- #
    linha("7) Conta-alvo configurada no app")
    alvo = "giulliano.luccia@jpmorgan.com"
    try:
        import json
        import os
        cfg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
        with open(cfg_path, "r", encoding="utf-8") as fp:
            alvo = json.load(fp).get("conta_email", alvo)
    except Exception:
        pass
    print(f"  Conta-alvo: {alvo}")

    achou = any(alvo.lower() == c.lower() for c in contas)
    if achou:
        ok("A conta-alvo BATE com uma conta do perfil. Deve conectar.")
    else:
        falha("A conta-alvo NÃO bate (exatamente) com nenhuma conta acima.")
        info("CORREÇÃO: no painel, ajuste o campo 'Conta conectada' para um")
        info("dos endereços SMTP exatos listados na seção 5 (ex.: o domínio")
        info("real pode ser @jpmchase.com em vez de @jpmorgan.com).")
        info("Alternativa: deixe o campo VAZIO para usar a conta padrão.")

    pythoncom.CoUninitialize()
    linha("Diagnóstico concluído")
    print("  Copie TODA a saída acima para análise.\n")


if __name__ == "__main__":
    main()
