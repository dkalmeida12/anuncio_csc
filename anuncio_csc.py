import pandas as pd
from datetime import datetime, date
import re
import unicodedata
from difflib import SequenceMatcher
import streamlit as st
import io
import requests
from typing import Tuple, Dict, Optional, List


# =========================
# CONFIG: GOOGLE SHEETS (PÚBLICO)
# =========================
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/10izQWPLAk3nv46Pl7ShzchReY3SjZdDl9KgboGQMAWg/edit?usp=sharing"
SHEET_ID_PATTERN = re.compile(r"/spreadsheets/d/([a-zA-Z0-9-_]+)")


def extrair_sheet_id(url: str) -> str:
    """Extrai ID da planilha Google Sheets."""
    m = SHEET_ID_PATTERN.search(str(url))
    return m.group(1) if m else ""


def baixar_sheets_publico_xlsx(sheet_url: str) -> bytes:
    """Baixa planilha pública do Google Sheets."""
    sheet_id = extrair_sheet_id(sheet_url)
    if not sheet_id:
        raise ValueError("Não foi possível extrair o ID da planilha a partir da URL informada.")
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    r = requests.get(export_url, timeout=30)
    r.raise_for_status()
    return r.content


# =========================
# EFETIVO CSC-PM (INTEGRADO NO CÓDIGO)
# IMPORTANTE: os * devem ser mantidos (WhatsApp negrito)
# =========================
# COMPRAS,135.147-7,*2ºTEN*,QOC,*CLEUBER* Ferreira da Silva ======== movimentado para a DAL 06 em 26/01/26

EFETIVO_CSC = """SEÇÃO,NÚMERO,P  / G,QUADRO,NOME
CHEFE,126.554-5,*TEN CEL*,QOPM,*LEONARDO* de *CASTRO* Ferreira
SUBCHEFE,089.655-5,*MAJ*,QOR,Jorge Aparecido *GOMES*
LICITAÇÃO,161.300-9,*CAP*,QOPM,Thiago Fernandes *PALMEIRA*
LICITAÇÃO,100.433-2,*2ºTEN*,QOR,*CLAUDIO* Marcio da Silva
LICITAÇÃO,087.650-8,*SUBTEN*,QPR,Sérgio Bernardino de *SENA*
LICITAÇÃO,154.178-8,*2ºSGT*,QPPM,Thiago *LUIZ TEIXEIRA*
COMPRAS,134.166-8,*CAP*,QOPM,Samuel Luiz *VIEIRA*
COMPRAS,142.483-7,*Asp a Of*,QPEP, *VALTER* Martins da Silva
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
ALMOX,167.118-9,*ASPM*,CIVIL,*DANIELLE* Gomes Figueiroa
S PRODUÇÃO GRÁFICA,113.505-2,*1ºSGT*,QPR,Carlos *LAÉRCIO* de Souza
S PRODUÇÃO GRÁFICA,094.227-6,*2ºTEN*,QOR,Vilmo Gonçalves *LEMOS*
S MANUTENÇÃO,087.957-7,*2ºTEN*,QOR,Joaquim *ARAÚJO* de Oliveira
S MANUTENÇÃO,102.773-9,*2ºSGT*,QPR,*NIVAL* Neves de Carvalho
S MANUTENÇÃO,090.803-8,*2ºSGT*,QPR,Arnaldo *BENTO* Pereira
S MANUTENÇÃO,097.538-3,*2ºSGT*,QPR,Carlos R. *SANTIAGO* dos Santos
S MANUTENÇÃO,127.860-5,*3ºSGT*,QPPM,Wagner *VITOR* dos Santos
"""

# =========================
# CONSTANTES
# =========================
QUADRO_CATEGORIA = {
    "QOPM": "OFICIAIS", "QOR": "OFICIAIS", "QOC": "OFICIAIS",
    "QPR": "PRAÇAS", "QPPM": "PRAÇAS", "QPE": "PRAÇAS",
    "CIVIL": "CIVIS"
}

STATUS_KEYWORDS = [
    (["férias", "ferias"], 1),
    (["licença", "licenca"], 2),
    (["ausente"], 3),
    (["folga"], 4),
    (["dispensa"], 5),
    (["presente"], 6),
]

# Regex para tokens em negrito (WhatsApp)
STAR_TOKEN_PATTERN = re.compile(r"\*([^*]+)\*")

# Padrões para extrair nome do cabeçalho do formulário
POSTO_PATTERNS = [
    (re.compile(r'^[\s]*ASPM[\s]+', re.IGNORECASE), ''),
    (re.compile(r'^[\s]*\d+[º°][\s]*', re.IGNORECASE), ''),
    (re.compile(r'^[\s]*(TEN[\s]*CEL|MAJ|CAP|SUB[\s]*TENENTE|SUBTENENTE|TEN|SGT|CB)[\s]+', re.IGNORECASE), ''),
    (re.compile(r'^[\s]*\d+[º°]?(TEN|SGT)[\s]+', re.IGNORECASE), ''),
]

# Ranking hierárquico (mais antigo -> mais moderno)
RANK_OFICIAIS = {
    "TEN CEL": 10, "TENENTE CORONEL": 10,
    "MAJ": 20, "MAJOR": 20,
    "CAP": 30, "CAPITAO": 30, "CAPITÃO": 30,
    "1° TEN": 40, "1 TEN": 40, "1 TENENTE": 40, "PRIMEIRO TENENTE": 40,
    "2° TEN": 50, "2 TEN": 50, "2 TENENTE": 50, "SEGUNDO TENENTE": 50,
}

RANK_PRACAS = {
    "SUBTEN": 10, "SUB TEN": 10, "SUBTENENTE": 10,
    "1° SGT": 20, "1 SGT": 20, "1 SARGENTO": 20,
    "2° SGT": 30, "2 SGT": 30, "2 SARGENTO": 30,
    "3° SGT": 40, "3 SGT": 40, "3 SARGENTO": 40,
    "CB": 50, "CABO": 50,
    "SD": 60, "SOLDADO": 60,
}


# =========================
# SESSION STATE
# =========================
def init_session_state():
    """Inicializa variáveis do session state."""
    defaults = {
        "df_formulario": None,
        "fonte_ok": False,
        "periodos_aplicados": False,
        "periodos_inseridos": {},
        "periodos_memoria": {},
        "last_sheet_url": DEFAULT_SHEET_URL,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# =========================
# AUXILIARES (normalização / matching)
# =========================
def remover_asteriscos(s: str) -> str:
    """Remove asteriscos de uma string."""
    return s.replace("*", "") if s else ""


def remover_acentos(s: str) -> str:
    """Remove acentos de uma string."""
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))


@st.cache_data
def normalizar_nome(nome: str) -> str:
    """
    Normalização para matching:
    - remove *
    - upper
    - remove acentos
    - remove pontuação
    """
    if pd.isna(nome):
        return ""
    s = remover_asteriscos(str(nome)).strip().upper()
    s = remover_acentos(s)
    s = re.sub(r"[^A-Z\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def normalizar_posto_display(posto: str) -> str:
    """
    Mantém * (WhatsApp), apenas normaliza º->° e espaços.
    """
    s = str(posto).strip()
    s = s.replace("º", "°")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def extrair_nome_completo_da_coluna(nome_coluna: str) -> str:
    """Extrai nome completo do militar do cabeçalho da coluna."""
    s = str(nome_coluna).strip()

    idx = s.upper().rfind(" PM ")
    if idx != -1:
        return s[idx + 4:].strip()

    for pattern, repl in POSTO_PATTERNS:
        s = pattern.sub(repl, s)

    return s.strip()


def similaridade(a: str, b: str) -> float:
    """Calcula similaridade entre duas strings."""
    return SequenceMatcher(None, a, b).ratio()


def encontrar_militar(nome_extraido: str, efetivo_dict: Dict, limiar: float = 0.88) -> Tuple[Optional[str], Optional[Dict]]:
    """Encontra militar no dicionário usando matching exato ou por similaridade."""
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


# =========================
# EXIBIÇÃO: SOMENTE TOKENS ENTRE *...*
# =========================
def extrair_tokens_negrito(texto: str) -> List[str]:
    """
    Retorna somente os trechos entre *...*, preservando os asteriscos.
    Ex: "*LEONARDO* de *CASTRO*" -> ["*LEONARDO*", "*CASTRO*"]
    """
    if not texto:
        return []
    return [f"*{m.group(1).strip()}*" for m in STAR_TOKEN_PATTERN.finditer(str(texto)) if m.group(1).strip()]


def formatar_nome_posto_somente_negritos(dados: Dict) -> str:
    """
    Ex.: posto "*TEN CEL*" e nome "*LEONARDO* de *CASTRO* Ferreira"
      -> "*TEN CEL*, *LEONARDO* *CASTRO*"
    """
    posto = str(dados.get("posto_display", "")).strip()
    nome = str(dados.get("nome_display", "")).strip()

    posto_tokens = extrair_tokens_negrito(posto)
    nome_tokens = extrair_tokens_negrito(nome)

    posto_out = posto_tokens[0] if posto_tokens else posto
    nome_out = " ".join(nome_tokens) if nome_tokens else nome

    return f"{posto_out}, {nome_out}".strip()


# =========================
# STATUS / PERÍODOS
# =========================
def classificar_status(resp: str) -> Tuple[str, int]:
    """Classifica status da resposta e retorna prioridade."""
    resp_lower = str(resp).strip().lower()

    if resp_lower == "presente":
        return "Presente", 6
    if resp_lower == "ausente":
        return "Ausente", 3
    if resp_lower == "folga":
        return "Folga", 4
    if "dispensa" in resp_lower:
        return "Dispensa pela Chefia", 5

    for keywords, priority in STATUS_KEYWORDS:
        if any(kw in resp_lower for kw in keywords):
            return str(resp).strip(), priority

    return str(resp).strip(), 50


def precisa_periodo(status: str) -> bool:
    """Verifica se o status requer informação de período."""
    sl = str(status).lower()
    return ("férias" in sl or "ferias" in sl or "licença" in sl or "licenca" in sl)


def validar_periodo(inicio: date, fim: date) -> bool:
    """Valida se o período é consistente."""
    return fim >= inicio


def formatar_periodo(inicio: date, fim: date) -> str:
    """Formata período para exibição."""
    return f"{inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}"


def ordem_status(s: str) -> int:
    """Retorna ordem de prioridade para ordenação de status."""
    sl = str(s).lower()
    for keywords, priority in STATUS_KEYWORDS:
        if any(kw in sl for kw in keywords):
            return priority
    return 50


# =========================
# ORDENAÇÃO HIERÁRQUICA
# =========================
def limpar_para_ranking(texto: str) -> str:
    """Normaliza texto para comparação de ranking."""
    if not texto:
        return ""
    s = remover_asteriscos(str(texto)).upper().strip()
    s = remover_acentos(s)
    s = s.replace("º", "°")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def extrair_chave_posto(posto_display: str) -> str:
    """Extrai chave do posto para ranking."""
    s = limpar_para_ranking(posto_display)
    s = s.replace("1°TEN", "1° TEN").replace("2°TEN", "2° TEN").replace("3°TEN", "3° TEN")
    s = s.replace("1°SGT", "1° SGT").replace("2°SGT", "2° SGT").replace("3°SGT", "3° SGT")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def rank_hierarquico(dados_militar: Dict) -> int:
    """Retorna rank hierárquico do militar (menor = mais antigo)."""
    categoria = dados_militar.get("categoria", "")
    posto_display = dados_militar.get("posto_display", "")
    chave = extrair_chave_posto(posto_display)

    if categoria == "OFICIAIS":
        if chave in RANK_OFICIAIS:
            return RANK_OFICIAIS[chave]
        for k, v in RANK_OFICIAIS.items():
            if k in chave:
                return v
        return 900

    if categoria == "PRAÇAS":
        if chave in RANK_PRACAS:
            return RANK_PRACAS[chave]
        for k, v in RANK_PRACAS.items():
            if k in chave:
                return v
        return 900

    return 999  # CIVIS / outros


# =========================
# CONVERSÃO SEGURA DE DATAS
# =========================
def to_datetime_safe(series: pd.Series) -> pd.Series:
    """
    Converte datas de forma segura:
    - aceita strings e números (Excel)
    - valores inválidos viram NaT (não quebram o app)
    """
    s = series.copy()

    # Se vier número do Excel (dias desde 1899-12-30), tenta converter
    s_num = pd.to_numeric(s, errors="coerce")
    s_dt_excel = pd.to_datetime(s_num, unit="D", origin="1899-12-30", errors="coerce")

    # Converte também como string (caso padrão)
    s_dt_str = pd.to_datetime(s, errors="coerce", dayfirst=True)

    # Prioriza o que deu certo (excel > string)
    out = s_dt_excel.combine_first(s_dt_str)

    return out


# =========================
# PROCESSAMENTO DE DADOS
# =========================
@st.cache_data
def carregar_efetivo() -> Dict:
    """Carrega e processa dados do efetivo (cached)."""
    df_efetivo = pd.read_csv(io.StringIO(EFETIVO_CSC))

    for col in ["SEÇÃO", "NÚMERO", "P  / G", "QUADRO", "NOME"]:
        df_efetivo[col] = df_efetivo[col].astype(str).str.strip()

    efetivo_dict = {}

    for _, row in df_efetivo.iterrows():
        quadro = row["QUADRO"].upper()
        categoria = QUADRO_CATEGORIA.get(quadro)
        if not categoria:
            continue

        nome_display = row["NOME"]
        posto_display = normalizar_posto_display(row["P  / G"])
        nome_norm = normalizar_nome(nome_display)

        efetivo_dict[nome_norm] = {
            "categoria": categoria,
            "posto_display": posto_display,
            "nome_display": nome_display,
            "quadro": quadro,
            "secao": row["SEÇÃO"].upper().strip(),
        }

    return efetivo_dict


def processar_respostas(df_hoje: pd.DataFrame, efetivo_dict: Dict) -> Dict:
    """Processa respostas do formulário e retorna dicionário de status."""
    respostas_dict = {}
    secoes_processadas = set()
    colunas_militares = df_hoje.columns[4:]

    for _, row in df_hoje.iterrows():
        secao = str(row["Seção:"])
        if secao in secoes_processadas:
            continue
        secoes_processadas.add(secao)

        for col in colunas_militares:
            valor = row[col]
            if pd.isna(valor) or str(valor).strip() == "":
                continue

            nome_militar = extrair_nome_completo_da_coluna(str(col).strip())
            chave_efetivo, militar_encontrado = encontrar_militar(nome_militar, efetivo_dict)

            if not militar_encontrado:
                continue

            respostas = [r.strip() for r in str(valor).strip().split(",") if r.strip()]
            candidatos = [classificar_status(resp) for resp in respostas]

            if candidatos:
                status_texto_exato = min(candidatos, key=lambda x: x[1])[0]
                respostas_dict[chave_efetivo] = {
                    "status": status_texto_exato,
                    "dados": militar_encontrado,
                }

    return respostas_dict


def organizar_categorias(
    efetivo_dict: Dict,
    respostas_dict: Dict,
    periodos_inseridos: Dict
) -> Tuple[Dict, Dict, List[str]]:
    """
    Organiza dados por categoria com ordenação hierárquica.
    Guarda listas como tuplas (rank, texto) para ordenar depois.
    """
    categorias_dados = {
        cat: {"presentes": [], "afastamentos": {}, "total": 0}
        for cat in ["OFICIAIS", "PRAÇAS", "CIVIS"]
    }

    faltantes_por_secao = {}
    militares_nao_informados = []

    for nome_norm, dados in efetivo_dict.items():
        categoria = dados["categoria"]
        categorias_dados[categoria]["total"] += 1

        resposta = respostas_dict.get(nome_norm)
        if not resposta:
            secao = dados.get("secao", "SEM SEÇÃO")
            faltantes_por_secao[secao] = faltantes_por_secao.get(secao, 0) + 1
            militares_nao_informados.append(f"{formatar_nome_posto_somente_negritos(dados)} ({secao})")
            continue

        status = str(resposta["status"]).strip()

        # TEXTO de exibição: SOMENTE os tokens em *
        posto_nome_display = formatar_nome_posto_somente_negritos(dados)
        r = rank_hierarquico(dados)

        # período (quando aplicável)
        if precisa_periodo(status) and nome_norm in periodos_inseridos:
            ini, fim = periodos_inseridos[nome_norm]
            posto_nome_saida = f"{posto_nome_display} - {formatar_periodo(ini, fim)}"
        else:
            posto_nome_saida = posto_nome_display

        # presente x afastamento
        if "presente" in status.lower() or status == "Presente":
            categorias_dados[categoria]["presentes"].append((r, posto_nome_display))
        else:
            categorias_dados[categoria]["afastamentos"].setdefault(status, []).append((r, posto_nome_saida))

    return categorias_dados, faltantes_por_secao, militares_nao_informados


def gerar_anuncio(
    data_formatada: str,
    categorias_dados: Dict,
    faltantes_por_secao: Dict
) -> Tuple[str, int, int]:
    """Gera texto do anúncio de presença com ordenação hierárquica."""
    anuncio_parts = [
        "Sr. Cel DAL, bom dia!\n",
        
        "Anúncio CSC-PM",
        data_formatada,
        ""
    ]

    total_militares = 0
    total_civis = 0

    for categoria in ["OFICIAIS", "PRAÇAS", "CIVIS"]:
        dados_cat = categorias_dados[categoria]

        if categoria == "CIVIS":
            total_civis = len(dados_cat["presentes"])
        else:
            total_militares += len(dados_cat["presentes"])

        anuncio_parts.extend([
            f"*{categoria}*",
            "Efetivo total: ",
            f"🔸{dados_cat['total']} - CSC-PM",
            ""
        ])

        # Presentes (ordenar por hierarquia)
        if dados_cat["presentes"]:
            presentes_ordenados = sorted(dados_cat["presentes"], key=lambda x: (x[0], x[1]))
            anuncio_parts.append(f"🔹{len(presentes_ordenados)} Presentes:")
            anuncio_parts.extend(f"    {i}. {txt}" for i, (_, txt) in enumerate(presentes_ordenados, 1))
            anuncio_parts.append("")

        # Afastamentos por status (e por hierarquia dentro de cada status)
        for status in sorted(dados_cat["afastamentos"].keys(), key=ordem_status):
            lista = dados_cat["afastamentos"][status]
            lista_ordenada = sorted(lista, key=lambda x: (x[0], x[1]))
            anuncio_parts.append(f"🔹{len(lista_ordenada)} {status}")
            anuncio_parts.extend(f"    {i}. {txt}" for i, (_, txt) in enumerate(lista_ordenada, 1))
            anuncio_parts.append("")

        anuncio_parts.append("")

    # Seções sem resposta
    if faltantes_por_secao:
        itens = sorted(faltantes_por_secao.items(), key=lambda x: (-x[1], x[0]))
        anuncio_parts.append(f"❌ Seções sem resposta ({len(itens)}):")
        for secao, qtd in itens:
            anuncio_parts.append(f"➡️ {secao}({qtd} servidores no total);")
        anuncio_parts.append("")

    anuncio_parts.extend([
        "Anúncio passado:",
        "[PREENCHER MANUALMENTE]",
        "[PREENCHER HORA]",
        "➖➖➖➖➖ ➖ ➖",
        "*Efetivo presente*:",
        f"*Militares: {total_militares}*",
        f"*Civis: {total_civis}*"
    ])

    return "\n".join(anuncio_parts), total_militares, total_civis


# =========================
# UI PRINCIPAL (STREAMLIT)
# =========================
def main():
    """Função principal do aplicativo."""
    init_session_state()

    st.title("GERADOR DE ANÚNCIO DE PRESENÇA CSC-PM v4.0")
    st.markdown("---")

    # Sidebar: controles
    with st.sidebar:
        st.subheader("⚙️ Controles")

        if st.button("🔄 Limpar carregamento (reset)"):
            st.session_state.df_formulario = None
            st.session_state.fonte_ok = False
            st.session_state.periodos_aplicados = False
            st.session_state.periodos_inseridos = {}
            st.rerun()

        if st.button("🗑️ Limpar memória de períodos"):
            st.session_state.periodos_memoria = {}
            st.success("Memória de períodos limpa.")
            st.rerun()

        if st.button("⚡ Limpar cache (efetivo / normalização)"):
            st.cache_data.clear()
            st.success("Cache limpo. Recarregando...")
            st.rerun()

        st.markdown("---")
        st.caption("v4.0 - Melhorias: ordenação hierárquica, exibição otimizada, conversão de datas robusta")

    # 1) Carregar planilha do formulário
    st.subheader("1️⃣ Carregar planilha do formulário")

    modo = st.radio(
        "Como deseja carregar a planilha do formulário?",
        ["URL Google Sheets (público) - automático", "Upload (XLS/XLSX)"],
        horizontal=True
    )

    if modo == "URL Google Sheets (público) - automático":
        sheet_url = st.text_input("URL do Google Sheets (público)", value=st.session_state.last_sheet_url)

        if st.button("📥 Baixar planilha"):
            try:
                with st.spinner("Baixando planilha..."):
                    xlsx_bytes = baixar_sheets_publico_xlsx(sheet_url)
                    st.session_state.df_formulario = pd.read_excel(io.BytesIO(xlsx_bytes))
                    st.session_state.fonte_ok = True
                    st.session_state.periodos_aplicados = False
                    st.session_state.periodos_inseridos = {}
                    st.session_state.last_sheet_url = sheet_url
                st.success("✅ Planilha baixada e carregada com sucesso!")
            except Exception as e:
                st.error(f"❌ Erro ao baixar/ler a planilha: {e}")

    else:
        uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xls", "xlsx"])
        if uploaded_file is not None:
            try:
                st.session_state.df_formulario = pd.read_excel(uploaded_file)
                st.session_state.fonte_ok = True
                st.session_state.periodos_aplicados = False
                st.session_state.periodos_inseridos = {}
                st.success("✅ Planilha carregada via upload!")
            except Exception as e:
                st.error(f"❌ Erro ao ler a planilha: {e}")

    df_formulario = st.session_state.df_formulario
    if not st.session_state.fonte_ok or df_formulario is None:
        st.info("📌 Carregue a planilha para continuar.")
        st.stop()

    # 2) Leitura das respostas
    data_atual = datetime.now()
    data_formatada = data_atual.strftime("%d/%m/%Y")

    efetivo_dict = carregar_efetivo()

    # Validar colunas obrigatórias
    colunas_obrigatorias = {"Carimbo de data/hora", "Data do anúncio", "Seção:"}
    faltando = colunas_obrigatorias - set(df_formulario.columns.astype(str))
    if faltando:
        st.error(f"❌ A planilha não possui as colunas obrigatórias: {', '.join(sorted(faltando))}")
        st.stop()

    # Conversão segura de datas
    df_formulario["Carimbo de data/hora"] = to_datetime_safe(df_formulario["Carimbo de data/hora"])
    df_formulario["Data do anúncio"] = to_datetime_safe(df_formulario["Data do anúncio"])

    # Filtrar registros de hoje
    df_hoje = df_formulario[df_formulario["Data do anúncio"].dt.date == data_atual.date()].copy()

    st.markdown("---")
    st.subheader("2️⃣ Leitura das respostas")

    if df_hoje.empty:
        st.warning(f"⚠️ ATENÇÃO: Não há registros para a data {data_formatada}")
        st.info("Verifique se a 'Data do anúncio' no formulário corresponde à data de hoje.")
        st.stop()

    st.success(f"✅ Encontrados {len(df_hoje)} registro(s) para {data_formatada}")
    df_hoje = df_hoje.sort_values("Carimbo de data/hora", ascending=False)

    respostas_dict = processar_respostas(df_hoje, efetivo_dict)

    # 3) Períodos (Férias / Licença)
    afastados = [
        (chave, resp["dados"], resp["status"])
        for chave, resp in respostas_dict.items()
        if precisa_periodo(resp["status"])
    ]

    st.markdown("---")
    st.subheader("3️⃣ Informar períodos (Férias / Licença)")
    st.caption("No anúncio: `POSTO, NOMES EM *negrito* - dd/mm/aaaa a dd/mm/aaaa`")

    if afastados and not st.session_state.periodos_aplicados:
        st.write("Preencha início e fim e clique em **Aplicar períodos**.")

        with st.form("form_periodos"):
            novos_periodos = {}
            erros = []

            for chave_norm, dados, status in afastados:
                posto_nome_display = formatar_nome_posto_somente_negritos(dados)
                st.markdown(f"**{posto_nome_display}**  \n_{status}_")

                ini_padrao, fim_padrao = st.session_state.periodos_memoria.get(
                    chave_norm, (data_atual.date(), data_atual.date())
                )

                c1, c2 = st.columns(2)
                inicio = c1.date_input("Início", value=ini_padrao, key=f"ini_{chave_norm}")
                fim = c2.date_input("Fim", value=fim_padrao, key=f"fim_{chave_norm}")

                if not validar_periodo(inicio, fim):
                    erros.append(
                        f"{posto_nome_display}: fim ({fim.strftime('%d/%m/%Y')}) não pode ser anterior ao início ({inicio.strftime('%d/%m/%Y')})."
                    )

                novos_periodos[chave_norm] = (inicio, fim)
                st.markdown("---")

            submitted = st.form_submit_button("✅ Aplicar períodos")

        if submitted:
            if erros:
                st.error("❌ Corrija os períodos abaixo antes de prosseguir:")
                for e in erros:
                    st.write(f"• {e}")
                st.stop()

            st.session_state.periodos_inseridos = novos_periodos
            st.session_state.periodos_aplicados = True
            st.session_state.periodos_memoria.update(novos_periodos)
            st.rerun()

    elif not afastados:
        st.info("Nenhum militar com status de férias/licença nesta data.")
        st.session_state.periodos_aplicados = True

    periodos_inseridos = st.session_state.periodos_inseridos if st.session_state.periodos_aplicados else {}

    # Organizar e gerar anúncio (com ordem hierárquica)
    categorias_dados, faltantes_por_secao, militares_nao_informados = organizar_categorias(
        efetivo_dict, respostas_dict, periodos_inseridos
    )

    anuncio, _, _ = gerar_anuncio(data_formatada, categorias_dados, faltantes_por_secao)

    st.markdown("---")
    st.subheader("📢 ANÚNCIO GERADO:")
    st.code(anuncio, language="text")

    if faltantes_por_secao:
        with st.expander("👥 Ver militares que não responderam (conferência)"):
            for item in sorted(militares_nao_informados):
                st.write(f"• {item}")

    st.download_button(
        label="💾 Baixar Anúncio de Presença",
        data=anuncio.encode("utf-8"),
        file_name=f"anuncio_presenca_{data_atual.strftime('%Y%m%d')}.txt",
        mime="text/plain"
    )

    st.success("✅ PROCESSO CONCLUÍDO!")


if __name__ == "__main__":
    main()
