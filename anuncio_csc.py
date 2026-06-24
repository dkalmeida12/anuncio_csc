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
# CONFIG
# =========================
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/10izQWPLAk3nv46Pl7ShzchReY3SjZdDl9KgboGQMAWg/edit?usp=sharing"
SHEET_ID_PATTERN  = re.compile(r"/spreadsheets/d/([a-zA-Z0-9-_]+)")

# Nome exato da aba de efetivo na planilha Google Sheets
ABA_EFETIVO    = "EFETIVO CSC"
# Nome exato da aba de respostas do formulário
ABA_FORMULARIO = "Respostas ao formulário 1"


# =========================
# CONSTANTES
# =========================
QUADRO_CATEGORIA = {
    "QOPM": "OFICIAIS", "QOR": "OFICIAIS", "QOC": "OFICIAIS",
    "QPEP": "OFICIAIS",
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

STAR_TOKEN_PATTERN = re.compile(r"\*([^*]+)\*")

POSTO_PATTERNS = [
    (re.compile(r'^[\s]*ASPM[\s]+',                re.IGNORECASE), ''),
    (re.compile(r'^[\s]*Asp[\s]+a[\s]+Of[\s]+',    re.IGNORECASE), ''),
    (re.compile(r'^[\s]*\d+[º°][\s]*',             re.IGNORECASE), ''),
    (re.compile(r'^[\s]*(TEN[\s]*CEL|MAJ|CAP|SUB[\s]*TENENTE|SUBTENENTE|TEN|SGT|CB)[\s]+',
                re.IGNORECASE), ''),
    (re.compile(r'^[\s]*\d+[º°]?(TEN|SGT)[\s]+',  re.IGNORECASE), ''),
]

RANK_OFICIAIS = {
    "TEN CEL": 10, "TENENTE CORONEL": 10,
    "MAJ": 20,     "MAJOR": 20,
    "CAP": 30,     "CAPITAO": 30,
    "1° TEN": 40,  "1 TEN": 40,  "PRIMEIRO TENENTE": 40,
    "2° TEN": 50,  "2 TEN": 50,  "SEGUNDO TENENTE": 50,
    "ASP A OF": 60, "ASP": 60,
}

RANK_PRACAS = {
    "SUBTEN": 10,  "SUB TEN": 10,  "SUBTENENTE": 10,
    "1° SGT": 20,  "1 SGT": 20,    "1 SARGENTO": 20,
    "2° SGT": 30,  "2 SGT": 30,    "2 SARGENTO": 30,
    "3° SGT": 40,  "3 SGT": 40,    "3 SARGENTO": 40,
    "CB": 50,      "CABO": 50,
    "SD": 60,      "SOLDADO": 60,
}


# =========================
# SESSION STATE
# =========================
def init_session_state():
    defaults = {
        "df_formulario":     None,
        "df_efetivo_raw":    None,
        "fonte_ok":          False,
        "periodos_aplicados": False,
        "periodos_inseridos": {},
        "periodos_memoria":  {},
        "last_sheet_url":    DEFAULT_SHEET_URL,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# =========================
# GOOGLE SHEETS
# =========================
def extrair_sheet_id(url: str) -> str:
    m = SHEET_ID_PATTERN.search(str(url))
    return m.group(1) if m else ""


def baixar_aba_xlsx(sheet_id: str, nome_aba: str) -> bytes:
    """Baixa uma aba específica da planilha via export."""
    # Primeiro precisamos descobrir o gid da aba
    url_html = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    r = requests.get(url_html, timeout=30)
    r.raise_for_status()
    return r.content


def baixar_planilha_completa(sheet_url: str) -> Dict[str, pd.DataFrame]:
    """
    Baixa a planilha completa e retorna dict {nome_aba: DataFrame}.
    """
    sheet_id = extrair_sheet_id(sheet_url)
    if not sheet_id:
        raise ValueError("Não foi possível extrair o ID da planilha.")

    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    r = requests.get(export_url, timeout=30)
    r.raise_for_status()

    xlsx = pd.ExcelFile(io.BytesIO(r.content))
    abas = {}
    for nome in xlsx.sheet_names:
        abas[nome] = xlsx.parse(nome)
    return abas


# =========================
# AUXILIARES
# =========================
def remover_asteriscos(s: str) -> str:
    return s.replace("*", "") if s else ""


def remover_acentos(s: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(ch)
    )


def normalizar_nome(nome: str) -> str:
    if pd.isna(nome):
        return ""
    s = remover_asteriscos(str(nome)).strip().upper()
    s = remover_acentos(s)
    s = re.sub(r"[^A-Z\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def normalizar_posto_display(posto: str) -> str:
    s = str(posto).strip().replace("º", "°")
    return re.sub(r"\s+", " ", s).strip()


def extrair_nome_completo_da_coluna(nome_coluna: str) -> str:
    s   = str(nome_coluna).strip()
    idx = s.upper().rfind(" PM ")
    if idx != -1:
        return s[idx + 4:].strip()
    for pattern, repl in POSTO_PATTERNS:
        s = pattern.sub(repl, s)
    return s.strip()


def similaridade(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def encontrar_militar(
    nome_extraido: str,
    efetivo_dict: Dict,
    limiar: float = 0.88
) -> Tuple[Optional[str], Optional[Dict]]:
    nome_norm = normalizar_nome(nome_extraido)
    if nome_norm in efetivo_dict:
        return nome_norm, efetivo_dict[nome_norm]

    melhor_key, melhor_score = None, 0.0
    for key in efetivo_dict:
        sc = similaridade(nome_norm, key)
        if sc > melhor_score:
            melhor_score = sc
            melhor_key   = key

    if melhor_key and melhor_score >= limiar:
        return melhor_key, efetivo_dict[melhor_key]
    return None, None


# =========================
# EXIBIÇÃO
# =========================
def extrair_tokens_negrito(texto: str) -> List[str]:
    if not texto:
        return []
    return [
        f"*{m.group(1).strip()}*"
        for m in STAR_TOKEN_PATTERN.finditer(str(texto))
        if m.group(1).strip()
    ]


def formatar_nome_posto_somente_negritos(dados: Dict) -> str:
    posto_tokens = extrair_tokens_negrito(str(dados.get("posto_display", "")))
    nome_tokens  = extrair_tokens_negrito(str(dados.get("nome_display",  "")))
    posto_out    = posto_tokens[0] if posto_tokens else dados.get("posto_display", "")
    nome_out     = " ".join(nome_tokens) if nome_tokens else dados.get("nome_display", "")
    return f"{posto_out}, {nome_out}".strip()


# =========================
# STATUS / PERÍODOS
# =========================
def classificar_status(resp: str) -> Tuple[str, int]:
    rl = str(resp).strip().lower()
    if rl == "presente":           return "Presente", 6
    if rl == "ausente":            return "Ausente",  3
    if rl == "folga":              return "Folga",    4
    if "dispensa" in rl:           return "Dispensa pela Chefia", 5
    for kws, pri in STATUS_KEYWORDS:
        if any(k in rl for k in kws):
            return str(resp).strip(), pri
    return str(resp).strip(), 50


def precisa_periodo(status: str) -> bool:
    sl = str(status).lower()
    return "férias" in sl or "ferias" in sl or "licença" in sl or "licenca" in sl


def validar_periodo(inicio: date, fim: date) -> bool:
    return fim >= inicio


def formatar_periodo(inicio: date, fim: date) -> str:
    return f"{inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}"


def ordem_status(s: str) -> int:
    sl = str(s).lower()
    for kws, pri in STATUS_KEYWORDS:
        if any(k in sl for k in kws):
            return pri
    return 50


# =========================
# RANKING HIERÁRQUICO
# =========================
def limpar_para_ranking(texto: str) -> str:
    s = remover_asteriscos(str(texto)).upper().strip()
    s = remover_acentos(s)
    s = s.replace("º", "°")
    return re.sub(r"\s+", " ", s).strip()


def rank_hierarquico(dados: Dict) -> int:
    categoria = dados.get("categoria", "")
    chave     = limpar_para_ranking(dados.get("posto_display", ""))
    chave     = re.sub(r"(\d+)°(TEN|SGT)", r"\1° \2", chave)

    tabela = RANK_OFICIAIS if categoria == "OFICIAIS" else (
             RANK_PRACAS   if categoria == "PRAÇAS"   else {})

    if chave in tabela:
        return tabela[chave]
    for k, v in tabela.items():
        if k in chave:
            return v
    return 999 if categoria == "CIVIS" else 900


# =========================
# CONVERSÃO DE DATAS
# =========================
def to_datetime_safe(series: pd.Series) -> pd.Series:
    s_num   = pd.to_numeric(series, errors="coerce")
    s_excel = pd.to_datetime(s_num, unit="D", origin="1899-12-30", errors="coerce")
    s_str   = pd.to_datetime(series, errors="coerce", dayfirst=True)
    return s_excel.combine_first(s_str)


# =========================
# CARREGAR EFETIVO DO SHEETS
# =========================
def carregar_efetivo_do_df(df_raw: pd.DataFrame) -> Dict:
    """
    Lê o DataFrame da aba EFETIVO e monta o dicionário de militares.

    Formato esperado da aba (colunas obrigatórias):
        SEÇÃO | NÚMERO | P / G | QUADRO | NOME

    A coluna NOME aceita asteriscos para negrito WhatsApp, ex:
        *LEONARDO* de *CASTRO* Ferreira
    """
    # Normalizar nomes de colunas (remover espaços extras)
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Aceitar variações do cabeçalho "P  / G" ou "P / G"
    col_posto = next(
        (c for c in df.columns if re.match(r"P\s*/\s*G", c, re.IGNORECASE)), None
    )
    if col_posto is None:
        raise ValueError(
            "Coluna de posto/graduação não encontrada na aba de efetivo. "
            "Certifique-se de que existe uma coluna com cabeçalho 'P / G'."
        )

    colunas_necessarias = ["SEÇÃO", "NÚMERO", "QUADRO", "NOME", col_posto]
    for c in colunas_necessarias:
        if c not in df.columns:
            raise ValueError(f"Coluna obrigatória ausente na aba de efetivo: '{c}'")

    efetivo_dict = {}
    for _, row in df.iterrows():
        quadro    = str(row["QUADRO"]).strip().upper()
        categoria = QUADRO_CATEGORIA.get(quadro)
        if not categoria:
            continue  # linha em branco ou quadro desconhecido

        nome_display  = str(row["NOME"]).strip()
        posto_display = normalizar_posto_display(str(row[col_posto]))
        nome_norm     = normalizar_nome(nome_display)

        if not nome_norm:
            continue

        efetivo_dict[nome_norm] = {
            "categoria":     categoria,
            "posto_display": posto_display,
            "nome_display":  nome_display,
            "quadro":        quadro,
            "secao":         str(row["SEÇÃO"]).strip().upper(),
        }

    return efetivo_dict


# =========================
# PROCESSAMENTO
# =========================
def processar_respostas(df_hoje: pd.DataFrame, efetivo_dict: Dict) -> Dict:
    respostas_dict     = {}
    secoes_processadas = set()

    for _, row in df_hoje.iterrows():
        secao = str(row["Seção:"])
        if secao in secoes_processadas:
            continue
        secoes_processadas.add(secao)

        for col in df_hoje.columns[4:]:
            valor = row[col]
            if pd.isna(valor) or str(valor).strip() == "":
                continue

            nome_extraido     = extrair_nome_completo_da_coluna(str(col).strip())
            chave, encontrado = encontrar_militar(nome_extraido, efetivo_dict)
            if not encontrado:
                continue

            candidatos = [classificar_status(r.strip())
                          for r in str(valor).strip().split(",") if r.strip()]
            if candidatos:
                status = min(candidatos, key=lambda x: x[1])[0]
                respostas_dict[chave] = {"status": status, "dados": encontrado}

    return respostas_dict


def organizar_categorias(
    efetivo_dict:       Dict,
    respostas_dict:     Dict,
    periodos_inseridos: Dict
) -> Tuple[Dict, Dict, List[str]]:
    categorias_dados = {
        cat: {"presentes": [], "afastamentos": {}, "total": 0}
        for cat in ["OFICIAIS", "PRAÇAS", "CIVIS"]
    }
    faltantes_por_secao      = {}
    militares_nao_informados = []

    for nome_norm, dados in efetivo_dict.items():
        categoria = dados["categoria"]
        categorias_dados[categoria]["total"] += 1

        resposta = respostas_dict.get(nome_norm)
        if not resposta:
            secao = dados.get("secao", "SEM SEÇÃO")
            faltantes_por_secao[secao] = faltantes_por_secao.get(secao, 0) + 1
            militares_nao_informados.append(
                f"{formatar_nome_posto_somente_negritos(dados)} ({secao})"
            )
            continue

        status    = str(resposta["status"]).strip()
        disp_base = formatar_nome_posto_somente_negritos(dados)
        rank      = rank_hierarquico(dados)

        if precisa_periodo(status) and nome_norm in periodos_inseridos:
            ini, fim = periodos_inseridos[nome_norm]
            disp = f"{disp_base} - {formatar_periodo(ini, fim)}"
        else:
            disp = disp_base

        if "presente" in status.lower():
            categorias_dados[categoria]["presentes"].append((rank, disp_base))
        else:
            categorias_dados[categoria]["afastamentos"].setdefault(status, []).append(
                (rank, disp)
            )

    return categorias_dados, faltantes_por_secao, militares_nao_informados


def gerar_anuncio(
    data_formatada:      str,
    categorias_dados:    Dict,
    faltantes_por_secao: Dict
) -> Tuple[str, int, int]:
    partes = ["Sr. Cel DAL, bom dia!\n", "Anúncio CSC-PM", data_formatada, ""]
    total_militares = total_civis = 0

    for categoria in ["OFICIAIS", "PRAÇAS", "CIVIS"]:
        d = categorias_dados[categoria]
        if categoria == "CIVIS":
            total_civis = len(d["presentes"])
        else:
            total_militares += len(d["presentes"])

        partes += [f"*{categoria}*", "Efetivo total: ", f"🔸{d['total']} - CSC-PM", ""]

        if d["presentes"]:
            presentes = sorted(d["presentes"], key=lambda x: (x[0], x[1]))
            partes.append(f"🔹{len(presentes)} Presentes:")
            partes += [f"    {i}. {t}" for i, (_, t) in enumerate(presentes, 1)]
            partes.append("")

        for status in sorted(d["afastamentos"], key=ordem_status):
            lista = sorted(d["afastamentos"][status], key=lambda x: (x[0], x[1]))
            partes.append(f"🔹{len(lista)} {status}")
            partes += [f"    {i}. {t}" for i, (_, t) in enumerate(lista, 1)]
            partes.append("")

        partes.append("")

    if faltantes_por_secao:
        itens = sorted(faltantes_por_secao.items(), key=lambda x: (-x[1], x[0]))
        partes.append(f"❌ Seções sem resposta ({len(itens)}):")
        for secao, qtd in itens:
            partes.append(f"➡️ {secao} ({qtd} servidores no total);")
        partes.append("")

    partes += [
        "Anúncio passado:",
        "[PREENCHER MANUALMENTE]",
        "[PREENCHER HORA]",
        "➖➖➖➖➖ ➖ ➖",
        "*Efetivo presente*:",
        f"*Militares: {total_militares}*",
        f"*Civis: {total_civis}*",
    ]
    return "\n".join(partes), total_militares, total_civis


# =========================
# UI PRINCIPAL
# =========================
def main():
    init_session_state()

    st.title("GERADOR DE ANÚNCIO DE PRESENÇA CSC-PM v5.0")
    st.markdown("---")

    with st.sidebar:
        st.subheader("⚙️ Controles")

        if st.button("🔄 Reset completo"):
            for k in ["df_formulario", "df_efetivo_raw", "fonte_ok",
                      "periodos_aplicados", "periodos_inseridos"]:
                st.session_state[k] = None if "df" in k else False if "ok" in k or "aplic" in k else {}
            st.rerun()

        if st.button("🗑️ Limpar memória de períodos"):
            st.session_state.periodos_memoria = {}
            st.success("Memória limpa.")
            st.rerun()

        st.markdown("---")

        st.markdown(
            "**📋 Formato da aba EFETIVO CSC:**\n\n"
            "| SEÇÃO | NÚMERO | P / G | QUADRO | NOME |\n"
            "|---|---|---|---|---|\n"
            "| P1 | 166.850-8 | *1º TEN* | QOPM | *DIEGO* Kukiyama de *ALMEIDA* |\n\n"
            "- **QUADRO aceitos:** QOPM, QOR, QOC, QPEP → Oficiais | "
            "QPR, QPPM, QPE → Praças | CIVIL → Civis\n"
            "- Use `*TEXTO*` para negrito no WhatsApp"
        )

        st.caption("v5.0 — Efetivo dinâmico via Google Sheets")

    # ── 1) Carregar planilha ──────────────────────────────────
    st.subheader("1️⃣ Carregar planilha")
    st.info(
        "A planilha deve ter **duas abas**:\n"
        f"- `{ABA_FORMULARIO}` — respostas do Google Forms\n"
        f"- `{ABA_EFETIVO}` — efetivo CSC (editável por você diretamente no Sheets)"
    )

    modo = st.radio(
        "Fonte dos dados:",
        ["URL Google Sheets (público) — automático", "Upload (XLS/XLSX)"],
        horizontal=True
    )

    if modo == "URL Google Sheets (público) — automático":
        sheet_url = st.text_input("URL do Google Sheets", value=st.session_state.last_sheet_url)

        if st.button("📥 Baixar planilha"):
            try:
                with st.spinner("Baixando planilha..."):
                    abas = baixar_planilha_completa(sheet_url)

                # Verificar abas necessárias
                abas_disponiveis = list(abas.keys())
                aba_form  = next((a for a in abas_disponiveis if ABA_FORMULARIO.lower() in a.lower()), None)
                aba_efet  = next((a for a in abas_disponiveis if ABA_EFETIVO.lower()   in a.lower()), None)

                if not aba_form:
                    st.error(
                        f"❌ Aba de formulário não encontrada.\n"
                        f"Abas disponíveis: {abas_disponiveis}\n"
                        f"Esperado: '{ABA_FORMULARIO}'"
                    )
                    st.stop()

                if not aba_efet:
                    st.error(
                        f"❌ Aba de efetivo não encontrada.\n"
                        f"Abas disponíveis: {abas_disponiveis}\n"
                        f"Esperado: '{ABA_EFETIVO}' — crie essa aba no Sheets."
                    )
                    st.stop()

                st.session_state.df_formulario      = abas[aba_form]
                st.session_state.df_efetivo_raw     = abas[aba_efet]
                st.session_state.fonte_ok           = True
                st.session_state.periodos_aplicados = False
                st.session_state.periodos_inseridos = {}
                st.session_state.last_sheet_url     = sheet_url
                st.success(f"✅ Planilha carregada! Abas lidas: '{aba_form}' e '{aba_efet}'")

            except Exception as e:
                st.error(f"❌ Erro: {e}")

    else:
        uploaded = st.file_uploader("Escolha o arquivo Excel (.xlsx)", type=["xls", "xlsx"])
        if uploaded:
            try:
                xlsx = pd.ExcelFile(uploaded)
                abas_disponiveis = xlsx.sheet_names

                aba_form = next((a for a in abas_disponiveis if ABA_FORMULARIO.lower() in a.lower()), None)
                aba_efet = next((a for a in abas_disponiveis if ABA_EFETIVO.lower()   in a.lower()), None)

                if not aba_form:
                    st.error(f"❌ Aba '{ABA_FORMULARIO}' não encontrada. Abas: {abas_disponiveis}")
                    st.stop()
                if not aba_efet:
                    st.error(f"❌ Aba '{ABA_EFETIVO}' não encontrada. Abas: {abas_disponiveis}")
                    st.stop()

                st.session_state.df_formulario      = xlsx.parse(aba_form)
                st.session_state.df_efetivo_raw     = xlsx.parse(aba_efet)
                st.session_state.fonte_ok           = True
                st.session_state.periodos_aplicados = False
                st.session_state.periodos_inseridos = {}
                st.success("✅ Planilha carregada via upload!")

            except Exception as e:
                st.error(f"❌ Erro: {e}")

    if not st.session_state.fonte_ok:
        st.stop()

    # ── 2) Carregar efetivo dinâmico ──────────────────────────
    st.markdown("---")
    st.subheader("2️⃣ Efetivo CSC")

    try:
        efetivo_dict = carregar_efetivo_do_df(st.session_state.df_efetivo_raw)
        total_efetivo = len(efetivo_dict)
        of  = sum(1 for d in efetivo_dict.values() if d["categoria"] == "OFICIAIS")
        pr  = sum(1 for d in efetivo_dict.values() if d["categoria"] == "PRAÇAS")
        civ = sum(1 for d in efetivo_dict.values() if d["categoria"] == "CIVIS")
        st.success(
            f"✅ Efetivo carregado: **{total_efetivo} servidores** "
            f"({of} oficiais | {pr} praças | {civ} civis)"
        )
    except Exception as e:
        st.error(f"❌ Erro ao processar aba de efetivo: {e}")
        st.stop()

    # ── 3) Leitura das respostas ──────────────────────────────
    st.markdown("---")
    st.subheader("3️⃣ Leitura das respostas")

    df_formulario = st.session_state.df_formulario
    data_atual    = datetime.now()
    data_formatada = data_atual.strftime("%d/%m/%Y")

    colunas_obrigatorias = {"Carimbo de data/hora", "Data do anúncio", "Seção:"}
    faltando = colunas_obrigatorias - set(df_formulario.columns.astype(str))
    if faltando:
        st.error(f"❌ Colunas obrigatórias ausentes na aba de formulário: {', '.join(sorted(faltando))}")
        st.stop()

    df_formulario["Carimbo de data/hora"] = to_datetime_safe(df_formulario["Carimbo de data/hora"])
    df_formulario["Data do anúncio"]      = to_datetime_safe(df_formulario["Data do anúncio"])

    df_hoje = df_formulario[
        df_formulario["Data do anúncio"].dt.date == data_atual.date()
    ].copy()

    if df_hoje.empty:
        st.warning(f"⚠️ Nenhuma resposta para {data_formatada}.")
        st.stop()

    st.success(f"✅ {len(df_hoje)} registro(s) para {data_formatada}")
    df_hoje = df_hoje.sort_values("Carimbo de data/hora", ascending=False)

    respostas_dict = processar_respostas(df_hoje, efetivo_dict)

    # ── 4) Períodos ───────────────────────────────────────────
    afastados = [
        (chave, resp["dados"], resp["status"])
        for chave, resp in respostas_dict.items()
        if precisa_periodo(resp["status"])
    ]

    st.markdown("---")
    st.subheader("4️⃣ Períodos de férias / licença")

    if afastados and not st.session_state.periodos_aplicados:
        with st.form("form_periodos"):
            novos_periodos, erros = {}, []
            for chave_norm, dados, status in afastados:
                posto_nome = formatar_nome_posto_somente_negritos(dados)
                st.markdown(f"**{posto_nome}** — _{status}_")
                ini_pad, fim_pad = st.session_state.periodos_memoria.get(
                    chave_norm, (data_atual.date(), data_atual.date())
                )
                c1, c2 = st.columns(2)
                inicio = c1.date_input("Início", value=ini_pad, key=f"ini_{chave_norm}")
                fim    = c2.date_input("Fim",    value=fim_pad, key=f"fim_{chave_norm}")
                if not validar_periodo(inicio, fim):
                    erros.append(f"{posto_nome}: fim anterior ao início.")
                novos_periodos[chave_norm] = (inicio, fim)
                st.markdown("---")

            if st.form_submit_button("✅ Aplicar períodos"):
                if erros:
                    for e in erros:
                        st.error(e)
                    st.stop()
                st.session_state.periodos_inseridos  = novos_periodos
                st.session_state.periodos_aplicados  = True
                st.session_state.periodos_memoria.update(novos_periodos)
                st.rerun()
    elif not afastados:
        st.info("Nenhum militar em férias/licença hoje.")
        st.session_state.periodos_aplicados = True

    periodos_inseridos = (
        st.session_state.periodos_inseridos
        if st.session_state.periodos_aplicados else {}
    )

    # ── 5) Anúncio ────────────────────────────────────────────
    categorias_dados, faltantes_por_secao, militares_nao_informados = organizar_categorias(
        efetivo_dict, respostas_dict, periodos_inseridos
    )
    anuncio, _, _ = gerar_anuncio(data_formatada, categorias_dados, faltantes_por_secao)

    st.markdown("---")
    st.subheader("📢 ANÚNCIO GERADO:")
    st.code(anuncio, language="text")

    if faltantes_por_secao:
        with st.expander("👥 Militares sem resposta"):
            for item in sorted(militares_nao_informados):
                st.write(f"• {item}")

    st.download_button(
        label="💾 Baixar Anúncio",
        data=anuncio.encode("utf-8"),
        file_name=f"anuncio_presenca_{data_atual.strftime('%Y%m%d')}.txt",
        mime="text/plain"
    )

    st.success("✅ PROCESSO CONCLUÍDO!")


if __name__ == "__main__":
    main()
