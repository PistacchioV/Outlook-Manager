# Traffic Control - Outlook Manager

Painel web local (Flask + Bootstrap 5) que varre o **Inbox do Outlook**
periodicamente, filtra e-mails por **palavras-chave** ou **remetentes-chave**,
agrupa por **tópico/thread** e gera um **resumo** curto de cada tópico.

## Estrutura

```
Outlook-Manager/
├── app.py                # Servidor Flask + rotas da API JSON
├── outlook_manager.py    # Lógica do Outlook (pywin32) + background worker
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

## Histórico de resumos e respostas sugeridas

Cada varredura **acumula** o histórico de cada tópico em `historico_topicos.json`
(não versionado — contém corpo de e-mails):

- **Mensagens únicas:** novas mensagens são anexadas com dedup por assinatura
  (remetente + data + tamanho). No Outlook real o `ReceivedTime` é estável, então
  a mesma mensagem nunca conta duas vezes.
- **Trail de resumos:** um novo snapshot de resumo (com data/hora) é gravado
  apenas quando o resumo **muda** — então o card mostra a evolução do tópico ao
  longo do tempo (colapsável em "Histórico de resumos").
- **Resposta sugerida:** gerada a partir do **histórico completo** do tópico e
  exibida em cada card, com botão "Copiar".

## Resumos e respostas

Por padrão é **100% local** (determinístico, nada sai da máquina). Há um backend
**opcional** de LLM (Gemini) que, se configurado, melhora o resumo/resposta — e
cai automaticamente no local em qualquer erro.

### Backend opcional Gemini (`genai_backend.py`)

Desligado por padrão. Só ativa se houver credenciais no ambiente — sem isso,
nada é enviado para fora. ⚠️ **Em ambiente corporativo, valide com TI/Compliance
antes de ativar** (envia conteúdo de e-mail para o Google). Use de preferência
**Vertex AI** num projeto GCP aprovado.

```bat
:: Opção A — Gemini API por chave
set GOOGLE_API_KEY=...
:: Opção B — Vertex AI (projeto corporativo aprovado)
set GOOGLE_GENAI_USE_VERTEXAI=1
set GOOGLE_CLOUD_PROJECT=seu-projeto
set GOOGLE_CLOUD_LOCATION=us-central1
:: desligar à força: set TC_USE_GENAI=0   |   trocar modelo: set TC_GENAI_MODEL=...
```

Quando ativo, o card mostra um selo **Gemini**; senão, usa o resumo local. A LLM
só é chamada quando há mensagem nova no tópico (cache por nº de mensagens).

### Processamento local (padrão e fallback)

Roda na própria máquina, de forma determinística, em `outlook_manager.py`:

- `gerar_resumo(texto)` — **sumarização extrativa**: pontua cada frase pela
  frequência de palavras relevantes (descartando stopwords PT-BR), com bônus
  para a 1ª frase e para frases com termos de ação (urgente, prazo, liquidação,
  confirmar…), e escolhe as melhores reordenadas pela posição original.
- `propor_resposta(assunto, historico_texto, destinatario)` — rascunho por
  regras: detecta urgência, prazos, perguntas e pedidos de confirmação no
  histórico completo e monta uma resposta cordial pronta para revisão.

Para ajustar a saída, edite as listas `_STOPWORDS_PT` / `_TERMOS_ACAO` (resumo)
e as frases/gatilhos dentro de `propor_resposta` — sem nenhuma dependência extra.

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
