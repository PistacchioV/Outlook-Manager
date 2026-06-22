# Traffic Control - Outlook Manager

Painel web local (Flask + Bootstrap 5) que varre o **Inbox do Outlook**
periodicamente, filtra e-mails por **palavras-chave** ou **remetentes-chave**,
agrupa por **tópico/thread** e gera um **resumo** curto de cada tópico.

## Estrutura

```
Outlook-Manager/
├── app.py                # Servidor Flask + rotas da API JSON
├── outlook_manager.py    # Lógica do Outlook (pywin32) + background worker
├── resumidor.py          # Motor de resumo 100% LOCAL (sem IA/rede)
├── requirements.txt
└── templates/
    └── index.html        # Interface (Bootstrap 5, tema light/dark, polling)
```

## Como rodar (Windows — produção)

```bat
git clone https://github.com/PistacchioV/Outlook-Manager.git
cd Outlook-Manager
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

python app.py
```

Acesse: http://127.0.0.1:5000

### Desenvolvimento em macOS / Linux

Fora do Windows o `pywin32` não existe e o app entra em **MODO SIMULADO**
(e-mails fictícios) — útil para mexer na interface. No macOS a porta 5000 é
ocupada pelo AirPlay Receiver, então use outra porta via `PORT`:

```bash
pip install Flask
PORT=5001 python3 app.py   # http://127.0.0.1:5001
```

## Plataforma

- **Windows + Outlook desktop:** funciona de verdade via `win32com.client`.
- **macOS / Linux:** o import do `pywin32` falha e o app entra em
  **`MODO_SIMULADO`**, servindo e-mails fictícios — ótimo para desenvolver a UI.
  A navbar mostra um badge amarelo "MODO SIMULADO" nesse caso.

## ⚠️ Ambiente corporativo (Windows) — pontos de atenção

1. **"Novo Outlook" NÃO suporta COM.** A automação via `win32com` só funciona
   com o **Outlook Clássico**. Se a máquina estiver no "Novo Outlook" (rollout
   comum por GPO), o `Dispatch` falha e o painel mostra um erro pedindo para
   desligar o botão **"Novo Outlook"**. O Clássico precisa estar instalado.
2. **Primeiro acesso pode exibir um alerta de segurança do Outlook** ("um
   programa está tentando acessar o Outlook"). Em máquinas gerenciadas isso é
   controlado por GPO / antivírus — pode ser necessário liberar com o TI.
3. **Persistência:** a configuração é salva em `config.json` (ao lado dos
   scripts) e sobrevive a reinícios. Não versione esse arquivo se contiver
   e-mails internos.
4. **Segurança do servidor:** roda em `127.0.0.1` (somente local) e com
   `debug=False` (o debugger do Werkzeug permitiria execução remota de código).
5. **Iniciar com o Windows:** para deixar sempre ativo, crie um atalho de
   `pythonw app.py` na pasta *Inicializar* ou uma Tarefa Agendada. `pythonw`
   roda sem janela de console.

## Concorrência COM (importante)

A varredura roda em uma *background thread*. Como o Outlook é exposto via COM,
a thread chama `pythoncom.CoInitialize()` ao iniciar e `pythoncom.CoUninitialize()`
ao encerrar (em `outlook_manager.py::_worker_loop`). O reloader do Flask é
desativado (`use_reloader=False`) para não duplicar a thread.

## Histórico acumulado

Cada varredura **acumula** as mensagens de cada tópico em `historico_topicos.json`
(não versionado — contém corpo de e-mails). Novas mensagens são anexadas com
dedup por assinatura (remetente + data + tamanho). No Outlook real o
`ReceivedTime` é estável, então a mesma mensagem nunca conta duas vezes. O resumo
é recalculado apenas quando o nº de mensagens do tópico muda (cache).

## Resumo da conversa — 100% LOCAL (`resumidor.py`)

O resumo é **inteiramente local e determinístico**: nada sai da máquina, nenhuma
chamada de rede, nenhuma IA externa, nenhuma credencial — adequado a DLP /
compliance. O mesmo conjunto de e-mails sempre produz exatamente o mesmo resumo.

Cada card mostra **só o resumo** da chain, em três partes:

- **Visão geral** — síntese de 1 linha: assunto, nº de mensagens e participantes.
- **Pontos** — as frases mais relevantes da conversa, atribuídas a quem falou,
  em ordem cronológica.
- **Pontos de atenção** — prazos/datas, valores e IDs, perguntas em aberto e
  pedidos, com etiqueta (`Prazo` / `Pedido` / `Pergunta` / `Número`).

Como o motor "raciocina" (em `resumidor.py`) — uma versão enxuta de
**TF-IDF + LexRank**, em Python puro:

1. **Limpa** cada mensagem — remove citações (`De:/From:/On … wrote:`),
   assinaturas e disclaimers jurídicos (rodapé padrão do banco).
2. **Segmenta** em frases respeitando abreviações (`Sr.`), decimais
   (`R$ 1.500,00`) e horas, e **pontua** cada frase combinando:
   *informatividade* (peso TF-IDF — palavras raras/específicas pesam mais que
   as onipresentes) + *centralidade* (quanto a frase representa o resto da
   thread, via similaridade de cosseno) + bônus para ação, perguntas,
   datas/valores e a 1ª frase de cada mensagem.
3. **Seleciona** os melhores pontos e **remove redundância** (frases quase iguais
   via similaridade de Jaccard).
4. **Extrai** os pontos de atenção por regex (datas, valores, IDs) e gatilhos.

Para ajustar a saída, edite no `resumidor.py` as listas `_STOPWORDS` /
`_TERMOS_ACAO` / `_PEDIDO` e as expressões `_RE_DATA` / `_RE_VALOR` / `_RE_ID`
— sem nenhuma dependência extra.

## Endpoints da API

| Método | Rota                  | Descrição                          |
|--------|-----------------------|------------------------------------|
| GET    | `/api/config`         | Configuração atual                 |
| POST   | `/api/palavras`       | Adiciona palavra-chave             |
| DELETE | `/api/palavras`       | Remove palavra-chave               |
| POST   | `/api/pessoas`        | Adiciona remetente-chave           |
| DELETE | `/api/pessoas`        | Remove remetente-chave             |
| POST   | `/api/intervalo`      | Ajusta intervalo (segundos)        |
| POST   | `/api/conta`          | Define a conta/mailbox lida (SMTP) |
| GET    | `/api/topicos`        | Tópicos agrupados (consumido no polling) |
| POST   | `/api/varrer-agora`   | Dispara varredura imediata         |

## Conta conectada

O app lê a Inbox da conta configurada em **"Conta conectada"** no painel
(padrão: `giulliano.luccia@jpmorgan.com`). Numa máquina corporativa com várias
contas/mailboxes, ele localiza a Account pelo endereço SMTP e usa o Inbox do
respectivo `DeliveryStore` — independente do idioma do Outlook. Campo vazio =
conta padrão do perfil. A configuração persiste em `config.json` (não versionado).
