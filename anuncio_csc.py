import pandas as pd
from datetime import datetime, date
import re
import unicodedata
from difflib import SequenceMatcher
import streamlit as st
import io
import requests

# =========================
# CONFIG: GOOGLE SHEETS (PÚBLICO)
# =========================
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/10izQWPLAk3nv46Pl7ShzchReY3SjZdDl9KgboGQMAWg/edit?usp=sharing"

def extrair_sheet_id(url: str) -> str:
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", str(url))
    return m.group(1) if m else ""

def baixar_sheets_publico_xlsx(sheet_url: str) -> bytes:
    sheet_id = extrair_sheet_id(sheet_url)
    if not sheet_id:
        raise ValueError("Não foi possível extrair o SHEET_ID da URL informada.")
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    r = requests.get(export_url, timeout=30)
    r.raise_for_status()
    return r.content

# =========================
# EFETIVO CSC-PM (INTEGRADO NO CÓDIGO) - VERSÃO "GITHUB"
# =========================
EFETIVO_CSC = """SEÇÃO,NÚMERO,P  / G,QUADRO,NOME
CHEFE,126.554-5,*TEN CEL*,QOPM,*LEONARDO* de *CASTRO* Ferreira
SUBCHEFE,089.655-5,*MAJ*,QOR,Jorge Aparacido *GOMES*
LICITAÇÃO,161.300-9,*CAP*,QOPM,Thiago Fernandes *PALMEIRA*
LICITAÇÃO,100.433-2,*2ºTEN*,QOR,*CLAUDIO* Marcio da Silva
LICITAÇÃO,087.650-8,*SUBTEN*,QPR,Sérgio Bernardino de *SENA*
LICITAÇÃO,154.178-8,*2ºSGT*,QPPM,Thiago *LUIZ TEIXEIRA*
COMPRAS,134.166-8,*CAP*,QOPM,Samuel Luiz *VIEIRA*
COMPRAS,135.147-7,*2ºTEN*,QOC,*CLEUBER* Ferreira da Silva
COMPRAS,147.720-7,*3º SGT*,QPE,Herbert Diogo Frade *GARBAZZA*
P1,166.850-8,*1º TEN*,QOPM,*DIEGO* Kukiyama de *ALMEIDA*
P1,087.768-8,*1ºSGT*,QPR,*GLAUCO* Almeida Braz
P1,094.907-3,*2ºSGT*,QPR,Alexandre Augusto *CORREA*
P1,140.204-9,*3ºSGT*,QPPM,*LEONARDO* Gomes da Costa
P1,144.105-4,*3ºSGT*,QPPM,Mauro *JACOB* de Gouveia Quirino
P1,181.220-5,*3ºSGT*,QPPM,*NÚBIA* Aparecida Ribeiro
P1,174.777-3,*CB*,QPPM,Ana *CLÁUDIA* Tavares Caetano
P1,167.318-5,*ASPM*,CIVIL,*MARA* Cristina Duarte Pereira
SOFI,149.668-6,*CAP*,QOPM,*DIOGO* da Silva Rosa
SOFI,134.606-3,*1ºTEN*,QOC,Valter *ADRIANO* dos Santos
SOFI,134.927-3,*3º SGT*,QPPM,*WALITON* Keliton da Cruz
SOFI,146.417-1,*3º SGT*,QPPM,*TIAGO* Henrique da Silva
SOFI,146.299-3,*3º SGT*,QPPM,*GUSTAVO* Guimarães Afeito
ALMOX,099.519-1,*2ºTEN*,QOR,Walmir Márcio da *CRUZ*
ALMOX,099.309-7,*1ºSGT*,QPR,*OMAIR* Celso dos Santos
ALMOX,113.505-2,*1ºSGT*,QPR,Carlos *LAÉRCIO* de Souza
ALMOX,167.118-9,*ASPM*,CIVIL,*DANIELLE* Gomes Figueiroa
S PRODUÇÃO GRÁFICA,094.227-6,*2ºTEN*,QOR,Vilmo Gonçalves *LEMOS*
S MANUTENÇÃO,087.957-7,*2ºTEN*,QOR,Joaquim *ARAÚJO* de Oliveira
S MANUTENÇÃO,102.773-9,*2ºSGT*,QPR,*NIVAL* Neves de Carvalho
S MANUTENÇÃO,090.803-8,*2ºSGT*,QPR,Arnaldo *BENTO* Pereira
S MANUTENÇÃO,097.538-3,*2ºSGT*,QPR,Carlos R. *SANTIAGO* dos Santos
S MANUTENÇÃO,127.860-5,*3ºSGT*,QPPM,Wagner *VITOR* dos Santos"""

# =========================
# Session State Init
# =========================
if "df_formulario" not in st.session_state:
    st.session_state.df_formulario = None
if "fonte_ok" not in st.session_state:
    st.session_state.fonte_ok = False
if "periodos_aplicados" not in st.session_state:
    st.session_state.periodos_aplicados = False
if "periodos_inseridos" not in st.session_state:
    st.session_state.periodos_inseridos = {}
# ✅ NOVO: memória de períodos anteriores para pré-preenchimento
if "periodos_memoria" not in st.session_state:
    st.session_state.periodos_memoria = {}  # chave_norm -> (inicio, fim)

# =========================
# Funções auxiliares
# =========================
def limpar_markdown_asteriscos(s: str) -> str:
    return re.sub(r"\*", "", str(s)) if s is not None else ""

def remover_acentos(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def normalizar_nome(nome):
    if pd.isna(nome):
        return ""
    s = str(nome).strip().upper()
    s = remover_acentos(s)
    s = re.sub(r"[^A-Z\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalizar_posto(posto):
    posto = str(posto).strip().upper()
    posto = posto.replace('º', '°')
    posto = posto.replace('1º', '1°').replace('2º', '2°').replace('3º', '3°')
    posto = posto.replace('1º', '1°').replace('2º', '2°').replace('3º', '3°')
    return posto

def extrair_nome_completo_da_coluna(nome_coluna: str) -> str:
    s = str(nome_coluna).strip()
    if " PM " in s.upper():
        idx = s.upper().rfind(" PM ")
        return s[idx + 4:].strip()
    s = re.sub(r'^[\s]*ASPM[\s]+', '', s, flags=re.IGNORECASE)
    s = re.sub(r'^[\s]*\d+[º°][\s]*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'^[\s]*(TEN[\s]*CEL|MAJ|CAP|SUB[\s]*TENENTE|SUBTENENTE|TEN|SGT|CB)[\s]+', '', s, flags=re.IGNORECASE)
    s = re.sub(r'^[\s]*\d+[º°]?(TEN|SGT)[\s]+', '', s, flags=re.IGNORECASE)
    return s.strip()

def similaridade(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()

def encontrar_militar(nome_extraido: str, efetivo_dict: dict, limiar: float = 0.88):
    nome_norm = normalizar_nome(nome_extraido)
    if nome_norm in efetivo_dict:
        return nome_norm, efetivo_dict[nome_norm]

    melhor_key = None
    melhor_score = 0.0
    for key in efetivo_dict.keys():
        sc = similaridade(nome_norm, key)
        if sc > melhor_score:
            melhor_score = sc
            melhor_key = key

    if melhor_key and melhor_score >= limiar:
        return melhor_key, efetivo_dict[melhor_key]
    return None, None

def prioridade_texto(resp_lower: str) -> int:
    if 'férias' in resp_lower or 'ferias' in resp_lower:
        return 1
    if 'licença' in resp_lower or 'licenca' in resp_lower:
        return 2
    if 'ausente' in resp_lower:
        return 3
    if 'folga' in resp_lower:
        return 4
    if 'dispensa' in resp_lower:
        return 5
    if 'presente' in resp_lower:
        return 6
    return 50

def ordem_status(s):
    sl = s.lower()
    if 'férias' in sl or 'ferias' in sl:
        return 1
    if 'licença' in sl or 'licenca' in sl:
        return 2
    if 'ausente' in sl:
        return 3
    if 'folga' in sl:
        return 4
    if 'dispensa' in sl:
        return 5
    return 50

def formatar_periodo(inicio: date, fim: date) -> str:
    return f"{inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}"

def precisa_periodo(status: str) -> bool:
    s = status.lower()
    return ('férias' in s or 'ferias' in s or 'licença' in s or 'licenca' in s)

# ✅ NOVO: validação simples
def validar_periodo(inicio: date, fim: date) -> bool:
    return fim >= inicio

# =========================
# UI
# =========================
st.title("GERADOR DE ANÚNCIO DE PRESENÇA CSC-PM v3.8")
st.markdown("---")

with st.sidebar:
    if st.button("Limpar carregamento (reset)"):
        st.session_state.df_formulario = None
        st.session_state.fonte_ok = False
        st.session_state.periodos_aplicados = False
        st.session_state.periodos_inseridos = {}
        # mantém memória por padrão (se quiser limpar também, descomente)
        # st.session_state.periodos_memoria = {}
        st.rerun()

    if st.button("Limpar memória de períodos"):
        st.session_state.periodos_memoria = {}
        st.success("Memória de períodos limpa.")
        st.rerun()

st.subheader("1) Carregar planilha do formulário")

modo = st.radio(
    "Como deseja carregar a planilha do formulário?",
    ["URL Google Sheets (público) - automático", "Upload (XLS/XLSX)"],
    horizontal=True
)

if modo == "URL Google Sheets (público) - automático":
    sheet_url = st.text_input("URL do Google Sheets (público)", value=DEFAULT_SHEET_URL)
    if st.button("Baixar planilha"):
        try:
            xlsx_bytes = baixar_sheets_publico_xlsx(sheet_url)
            st.session_state.df_formulario = pd.read_excel(io.BytesIO(xlsx_bytes))
            st.session_state.fonte_ok = True
            st.session_state.periodos_aplicados = False
            st.session_state.periodos_inseridos = {}
            st.success("✅ Planilha baixada e carregada com sucesso!")
        except Exception as e:
            st.error(f"❌ Erro ao baixar/ler a planilha: {e}")
else:
    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xls", "xlsx"])
    if uploaded_file is not None:
        st.session_state.df_formulario = pd.read_excel(uploaded_file)
        st.session_state.fonte_ok = True
        st.session_state.periodos_aplicados = False
        st.session_state.periodos_inseridos = {}
        st.success("✅ Planilha carregada via upload!")

df_formulario = st.session_state.df_formulario

if not st.session_state.fonte_ok or df_formulario is None:
    st.info("Carregue a planilha para continuar.")
    st.stop()

# =========================
# PROCESSAMENTO
# =========================
data_atual = datetime.now()
data_formatada = data_atual.strftime("%d/%m/%Y")

df_efetivo = pd.read_csv(pd.io.common.StringIO(EFETIVO_CSC))
for col in ['SEÇÃO', 'NÚMERO', 'P  / G', 'QUADRO', 'NOME']:
    if col in df_efetivo.columns:
        df_efetivo[col] = df_efetivo[col].apply(limpar_markdown_asteriscos).str.strip()

efetivo_dict = {}
for _, row in df_efetivo.iterrows():
    nome_completo = str(row['NOME']).strip()
    quadro = str(row['QUADRO']).strip().upper()
    posto_grad = normalizar_posto(str(row['P  / G']))
    secao_efetivo = str(row['SEÇÃO']).strip().upper()

    if quadro in ['QOPM', 'QOR', 'QOC']:
        categoria = 'OFICIAIS'
    elif quadro in ['QPR', 'QPPM', 'QPE']:
        categoria = 'PRAÇAS'
    elif quadro == 'CIVIL':
        categoria = 'CIVIS'
    else:
        categoria = None

    if categoria:
        nome_norm = normalizar_nome(nome_completo)
        efetivo_dict[nome_norm] = {
            'categoria': categoria,
            'posto_grad': posto_grad,
            'nome_completo': nome_completo,
            'quadro': quadro,
            'secao': secao_efetivo
        }

colunas_obrigatorias = {'Carimbo de data/hora', 'Data do anúncio', 'Seção:'}
faltando = colunas_obrigatorias - set(df_formulario.columns.astype(str))
if faltando:
    st.error(f"❌ A planilha não possui as colunas obrigatórias: {', '.join(sorted(faltando))}")
    st.stop()

df_formulario['Carimbo de data/hora'] = pd.to_datetime(df_formulario['Carimbo de data/hora'])
df_formulario['Data do anúncio'] = pd.to_datetime(df_formulario['Data do anúncio'])
df_hoje = df_formulario[df_formulario['Data do anúncio'].dt.date == data_atual.date()].copy()

st.markdown("---")
st.subheader("2) Leitura das respostas")

if df_hoje.empty:
    st.warning(f"⚠️ ATENÇÃO: Não há registros para a data {data_formatada}")
    st.info("Verifique se a 'Data do anúncio' no formulário corresponde à data de hoje.")
    st.stop()
else:
    st.success(f"✅ Encontrados {len(df_hoje)} registro(s) para {data_formatada}")

df_hoje = df_hoje.sort_values('Carimbo de data/hora', ascending=False)

respostas_dict = {}
secoes_processadas = set()
colunas_militares = df_formulario.columns[4:]

for _, row in df_hoje.iterrows():
    secao = str(row['Seção:'])
    if secao in secoes_processadas:
        continue
    secoes_processadas.add(secao)

    for col in colunas_militares:
        valor = row[col]
        if pd.isna(valor) or str(valor).strip() == '':
            continue

        nome_militar = extrair_nome_completo_da_coluna(str(col).strip())
        chave_efetivo, militar_encontrado = encontrar_militar(nome_militar, efetivo_dict, limiar=0.88)
        if not militar_encontrado:
            continue

        respostas = [r.strip() for r in str(valor).strip().split(',') if r.strip()]
        candidatos = []

        for resp in respostas:
            resp_lower = resp.lower()
            if 'presente' in resp_lower:
                candidatos.append(("Presente", 6))
            elif 'ausente' in resp_lower:
                candidatos.append(("Ausente", 3))
            elif 'folga' in resp_lower:
                candidatos.append(("Folga", 4))
            elif 'dispensa' in resp_lower:
                candidatos.append(("Dispensa pela Chefia", 5))
            elif 'férias' in resp_lower or 'ferias' in resp_lower:
                candidatos.append((resp, 1))
            elif 'licença' in resp_lower or 'licenca' in resp_lower:
                candidatos.append((resp, 2))
            else:
                candidatos.append((resp, prioridade_texto(resp_lower)))

        if not candidatos:
            continue

        candidatos.sort(key=lambda x: x[1])
        status_texto_exato = candidatos[0][0]
        respostas_dict[chave_efetivo] = {'status': status_texto_exato, 'dados': militar_encontrado}

# =========================
# 3) Períodos (Férias / Licença) com:
#    (A) Validação fim >= início
#    (B) Pré-preenchimento com memória
# =========================
afastados = []
for chave_norm, resp in respostas_dict.items():
    status = str(resp['status']).strip()
    if precisa_periodo(status):
        afastados.append((chave_norm, resp['dados'], status))

st.markdown("---")
st.subheader("3) Informar períodos (Férias / Licença)")

if afastados and not st.session_state.periodos_aplicados:
    st.write("Preencha início e fim e clique em **Aplicar períodos**. O anúncio será gerado em seguida.")

    with st.form("form_periodos"):
        novos_periodos = {}
        erros = []

        for chave_norm, dados, status in afastados:
            posto_nome = f"{dados['posto_grad']} {dados['nome_completo']}"
            st.markdown(f"**{posto_nome}**  \n_{status}_")

            # ✅ Pré-preenchimento: usa memória se existir; senão data atual
            if chave_norm in st.session_state.periodos_memoria:
                ini_padrao, fim_padrao = st.session_state.periodos_memoria[chave_norm]
            else:
                ini_padrao, fim_padrao = data_atual.date(), data_atual.date()

            c1, c2 = st.columns(2)
            inicio = c1.date_input("Início", value=ini_padrao, key=f"ini_{chave_norm}")
            fim = c2.date_input("Fim", value=fim_padrao, key=f"fim_{chave_norm}")

            # ✅ Validação: fim não pode ser menor que início
            if not validar_periodo(inicio, fim):
                erros.append(f"{posto_nome}: o fim ({fim.strftime('%d/%m/%Y')}) não pode ser anterior ao início ({inicio.strftime('%d/%m/%Y')}).")

            novos_periodos[chave_norm] = (inicio, fim)
            st.markdown("---")

        submitted = st.form_submit_button("Aplicar períodos")

    if submitted:
        if erros:
            st.error("❌ Corrija os períodos abaixo antes de prosseguir:")
            for e in erros:
                st.write(f"• {e}")
            st.stop()

        # salva períodos escolhidos
        st.session_state.periodos_inseridos = novos_periodos
        st.session_state.periodos_aplicados = True

        # ✅ grava na memória (para pré-preenchimento nas próximas execuções)
        for k, v in novos_periodos.items():
            st.session_state.periodos_memoria[k] = v

        st.rerun()

elif not afastados:
    st.info("Nenhum militar com status de férias/licença nesta data.")
    st.session_state.periodos_aplicados = True

periodos_inseridos = st.session_state.periodos_inseridos if st.session_state.periodos_aplicados else {}

# =========================
# Organizar por categoria / status + seções sem resposta
# =========================
categorias_dados = {
    'OFICIAIS': {'presentes': [], 'afastamentos': {}, 'total': 0},
    'PRAÇAS': {'presentes': [], 'afastamentos': {}, 'total': 0},
    'CIVIS': {'presentes': [], 'afastamentos': {}, 'total': 0}
}

faltantes_por_secao = {}
militares_nao_informados_nomes = []

for nome_norm, dados in efetivo_dict.items():
    categoria = dados['categoria']
    categorias_dados[categoria]['total'] += 1

    resposta = respostas_dict.get(nome_norm)
    if not resposta:
        secao = dados.get('secao', 'SEM SEÇÃO')
        faltantes_por_secao[secao] = faltantes_por_secao.get(secao, 0) + 1
        militares_nao_informados_nomes.append(f"{dados['posto_grad']} {dados['nome_completo']} ({secao})")
        continue

    status = str(resposta['status']).strip()
    posto_nome = f"{dados['posto_grad']} {dados['nome_completo']}"

    if precisa_periodo(status) and nome_norm in periodos_inseridos:
        ini, fim = periodos_inseridos[nome_norm]
        posto_nome_saida = f"{posto_nome} - {formatar_periodo(ini, fim)}"
    else:
        posto_nome_saida = posto_nome

    if 'presente' in status.lower() or status == "Presente":
        categorias_dados[categoria]['presentes'].append(posto_nome)
    else:
        categorias_dados[categoria]['afastamentos'].setdefault(status, []).append(posto_nome_saida)

# =========================
# Gerar anúncio
# =========================
anuncio = f"""Bom dia!
Segue anúncio do dia

Anúncio CSC-PM
{data_formatada}

"""

total_militares = 0
total_civis = 0

for categoria in ['OFICIAIS', 'PRAÇAS', 'CIVIS']:
    dados = categorias_dados[categoria]

    if categoria == 'CIVIS':
        total_civis = len(dados['presentes'])
    else:
        total_militares += len(dados['presentes'])

    anuncio += f"*{categoria}*\n"
    anuncio += "Efetivo total: \n"
    anuncio += f"🔸{dados['total']} - CSC-PM\n\n"

    if dados['presentes']:
        anuncio += f"🔹{len(dados['presentes'])} Presentes:\n"
        for idx, nome in enumerate(dados['presentes'], 1):
            anuncio += f"    {idx}. {nome}\n"
        anuncio += "\n"

    afast = dados['afastamentos']
    for status in sorted(afast.keys(), key=ordem_status):
        lista = afast[status]
        anuncio += f"🔹{len(lista)} {status}\n"
        for idx, info in enumerate(lista, 1):
            anuncio += f"    {idx}. {info}\n"
        anuncio += "\n"

    anuncio += "\n"

if faltantes_por_secao:
    itens = sorted(faltantes_por_secao.items(), key=lambda x: (-x[1], x[0]))
    total_secoes = len(itens)
    seta = "➡️"
    anuncio += f"❌ Seções sem resposta ({total_secoes}):\n"
    for secao, qtd in itens:
        anuncio += f"{seta} {secao}({qtd} servidores no total);\n"
    anuncio += "\n"

anuncio += f"""Anúncio passado:
[PREENCHER MANUALMENTE]
[PREENCHER HORA]
➖➖➖➖➖ ➖ ➖
*Efetivo presente*:
*Militares: {total_militares}*
*Civis: {total_civis}*"""

st.markdown("---")
st.subheader("📢 ANÚNCIO GERADO:")
st.code(anuncio, language='text')

if faltantes_por_secao:
    with st.expander("Ver militares que não responderam (conferência)"):
        for item in sorted(militares_nao_informados_nomes):
            st.write(f"• {item}")

st.download_button(
    label="Baixar Anúncio de Presença",
    data=anuncio.encode('utf-8'),
    file_name="anuncio_presenca.txt",
    mime="text/plain"
)

st.success("✅ PROCESSO CONCLUÍDO!")
