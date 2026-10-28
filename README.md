# anuncio_csc

Gerador do anúncio diário de presença do CSC-PM (Streamlit).

## Como o app acessa a planilha
O app lê e grava a planilha por meio de um Apps Script publicado como App da Web
(`apps_script/Code.gs`), protegido por token. A planilha não precisa ser pública.

Abas usadas:
- `Respostas ao formulário 1` — respostas do Google Forms
- `EFETIVO CSC` — SEÇÃO | NÚMERO | P / G | QUADRO | NOME
- `AFASTAMENTOS` — criada e mantida pelo app (MATRÍCULA | NOME | TIPO | DESCRIÇÃO | INÍCIO | FIM)

## Configuração
1. Na planilha: Extensões > Apps Script; cole `apps_script/Code.gs`.
2. Configurações do projeto > Propriedades do script: `TOKEN` = senha longa aleatória.
3. Implantar > Nova implantação > App da Web (Executar como: Eu; Acesso: Qualquer pessoa).
4. Streamlit Cloud > Settings > Secrets: ver `.streamlit/secrets.toml.example`.
