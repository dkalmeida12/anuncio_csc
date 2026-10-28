import json
import re
import unicodedata
from datetime import datetime, date
from difflib import SequenceMatcher
from typing import Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import streamlit as st


# =========================
# CONFIG
# =========================
VERSAO = "v6.0"
TZ = ZoneInfo("America/Sao_Paulo")

# Colunas da aba de respostas que NÃO são militares
PADROES_COLUNAS_META = (
    "carimbo de data", "data do anuncio", "secao", "e-mail", "email",
    "de quem ta prestando",
)

LIMIAR_SIMILARIDADE = 0.88
MARGEM_AMBIGUIDADE = 0.03


# =========================
# CONSTANTES
# =========================
QUADRO_CATEGORIA = {
    "QOPM": "OFICIAIS", "QOR": "OFICIAIS", "QOC": "OFICIAIS",
    "QPEP": "OFICIAIS",
    "QPR": "PRAÇAS", "QPPM": "PRAÇAS", "QPE": "PRAÇAS",
    "CIVIL": "CIVIS"
}

# Palavras-chave (sem acento) -> ordem de exibição no anúncio
STATUS_KEYWORDS = [
    (["ferias"], 1),
    (["licenca", "luto", "nojo"], 2),
    (["ausente"], 3),
    (["folga"], 4),
    (["dispensa"], 5),
    (["presente"], 6),
]

DISPENSA_CHEFIA_OPCOES = {"dispensa pela chefia", "dispensa da chefia", "dispensa chefia", "dispensa"}
OPCOES_PADRAO = {"presente", "ausente", "folga"} | DISPENSA_CHEFIA_OPCOES

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

# Quadros da ativa = 0 (aparecem primeiro); reconvocados (QOR/QPR) = 1 (aparecem depois)
QUADRO_RANK = {
    "QOPM": 0, "QOC": 0, "QPPM": 0, "QPE": 0, "QPEP": 0,
    "QOR":  1, "QPR":  1,
    "CIVIL": 0,
}


# =========================
# SESSION STATE
# =========================
def init_session_state():
    defaults = {
        "dados":              None,   # dict com DataFrames lidos da planilha
        "carregado_em":       None,
        "etapa_periodos":     "editar",   # editar | confirmar | aplicado
        "periodos_inseridos": {},
        "plano_gravacao":     [],
        "msg_gravacao":       None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def resetar_periodos():
    st.session_state.etapa_periodos = "editar"
    st.session_state.periodos_inseridos = {}
    st.session_state.plano_gravacao = []
    st.session_state.msg_gravacao = None


# =========================
# AUXILIARES
# =========================
def agora() -> datetime:
    return datetime.now(TZ)


def remover_asteriscos(s: str) -> str:
    return s.replace("*", "") if s else ""


def remover_acentos(s: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(ch)
    )


def sem_acento_lower(s: str) -> str:
    return remover_acentos(str(s).strip().lower())


def normalizar_nome(nome: str) -> str:
    if nome is None or (isinstance(nome, float) and pd.isna(nome)):
        return ""
    s = remover_asteriscos(str(nome)).strip().upper()
    s = remover_acentos(s)
    s = re.sub(r"[^A-Z\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def chave_matricula(mat: str) -> str:
    """Somente dígitos: '089.655-5' -> '0896555' (ignora sufixo interno '#n')."""
    return re.sub(r"\D", "", str(mat or "").split("#")[0])


def normalizar_posto_display(posto: str) -> str:
    s = str(posto).strip().replace("º", "°")
    return re.sub(r"\s+", " ", s).strip()


def texto_simples(s: str) -> str:
    return re.sub(r"\s+", " ", remover_asteriscos(str(s))).strip()


def extrair_nome_completo_da_coluna(nome_coluna: str) -> str:
    s   = str(nome_coluna).strip()
    idx = s.upper().rfind(" PM ")
    if idx != -1:
        return s[idx + 4:].strip()
    for pattern, repl in POSTO_PATTERNS:
        s = pattern.sub(repl, s)
    return s.strip()


def eh_coluna_meta(coluna: str) -> bool:
    c = sem_acento_lower(coluna)
    return any(p in c for p in PADROES_COLUNAS_META)


def similaridade(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def parse_data_br(valor) -> Optional[date]:
    s = str(valor or "").strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return None if pd.isna(dt) else dt.date()


def parse_datahora_br(valor) -> pd.Timestamp:
    s = str(valor or "").strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return pd.Timestamp(datetime.strptime(s, fmt))
        except ValueError:
            pass
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def fmt_data(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def valores_para_df(valores: List[List[str]]) -> pd.DataFrame:
    if not valores:
        return pd.DataFrame()
    cab = [str(h).strip() for h in valores[0]]
    n = len(cab)
    linhas = [(list(r) + [""] * n)[:n] for r in valores[1:]]
    return pd.DataFrame(linhas, columns=cab)


# =========================
# PONTE APPS SCRIPT
# =========================
def config_api() -> Tuple[str, str]:
    try:
        cfg = st.secrets["apps_script"]
        return str(cfg["url"]).strip(), str(cfg["token"]).strip()
    except Exception:
        st.error(
            "❌ Configuração ausente. Em **Settings → Secrets** do Streamlit Cloud, informe:\n\n"
            "```toml\n[apps_script]\nurl = \"https://script.google.com/macros/s/.../exec\"\n"
            "token = \"mesmo TOKEN definido no Apps Script\"\n```"
        )
        st.stop()


def chamar_api(acao: str, **params) -> Dict:
    url, token = config_api()
    corpo = json.dumps({"token": token, "acao": acao, **params}, ensure_ascii=False)
    r = requests.post(
        url, data=corpo.encode("utf-8"),
        headers={"Content-Type": "text/plain;charset=utf-8"}, timeout=60,
    )
    r.raise_for_status()
    try:
        resp = r.json()
    except ValueError:
        raise RuntimeError(
            "A ponte com a planilha não respondeu em JSON. Verifique se a implantação do "
            "Apps Script está com acesso 'Qualquer pessoa' e se a URL termina em /exec."
        )
    if not resp.get("ok"):
        raise RuntimeError(resp.get("erro", "Erro desconhecido na ponte com a planilha."))
    return resp


def carregar_dados(hoje: date) -> Dict:
    resp = chamar_api("ler", data=fmt_data(hoje))
    return {
        "formulario":   valores_para_df(resp.get("formulario", [])),
        "efetivo":      valores_para_df(resp.get("efetivo", [])),
        "afastamentos": preparar_afastamentos(resp.get("afastamentos", [])),
    }


# =========================
# EFETIVO
# =========================
def carregar_efetivo_do_df(df_raw: pd.DataFrame) -> Tuple[Dict, Dict[str, List[str]], List[str]]:
    """
    Retorna:
      efetivo    {chave (matrícula só dígitos): dados}
      nome_index {nome normalizado: [chaves]}
      alertas    lista de textos
    """
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    col_posto = next((c for c in df.columns if re.match(r"P\s*/\s*G", c, re.IGNORECASE)), None)
    if col_posto is None:
        raise ValueError("Coluna 'P / G' não encontrada na aba de efetivo.")
    for c in ["SEÇÃO", "NÚMERO", "QUADRO", "NOME"]:
        if c not in df.columns:
            raise ValueError(f"Coluna obrigatória ausente na aba de efetivo: '{c}'")

    efetivo, nome_index, alertas = {}, {}, []

    for _, row in df.iterrows():
        quadro    = str(row["QUADRO"]).strip().upper()
        categoria = QUADRO_CATEGORIA.get(quadro)
        nome_display = str(row["NOME"]).strip()
        if not categoria:
            if nome_display:
                alertas.append(f"Quadro desconhecido '{quadro}' para {texto_simples(nome_display)} — linha ignorada.")
            continue

        nome_norm = normalizar_nome(nome_display)
        if not nome_norm:
            continue

        mat_display = str(row["NÚMERO"]).strip()
        chave = chave_matricula(mat_display)
        if not chave:
            chave = f"SEM-MATRICULA:{nome_norm}"
            alertas.append(f"{texto_simples(nome_display)} está sem matrícula na aba de efetivo.")
        elif chave in efetivo:
            alertas.append(
                f"Matrícula {mat_display} repetida na aba de efetivo "
                f"({texto_simples(efetivo[chave]['nome_display'])} e {texto_simples(nome_display)})."
            )
            chave = f"{chave}#{len(efetivo)}"

        efetivo[chave] = {
            "chave":         chave,
            "matricula":     mat_display,
            "categoria":     categoria,
            "posto_display": normalizar_posto_display(str(row[col_posto])),
            "nome_display":  nome_display,
            "nome_norm":     nome_norm,
            "quadro":        quadro,
            "secao":         str(row["SEÇÃO"]).strip().upper(),
        }
        nome_index.setdefault(nome_norm, []).append(chave)

    for nome_norm, chaves in nome_index.items():
        if len(chaves) > 1:
            alertas.append(
                f"Nome repetido no efetivo: {nome_norm} ({len(chaves)} registros). "
                "Respostas para esse nome não serão associadas automaticamente."
            )

    return efetivo, nome_index, alertas


def encontrar_militar(nome_coluna: str, nome_index: Dict[str, List[str]]) -> Tuple[Optional[str], float, str]:
    """Retorna (chave, score, motivo). chave=None quando não reconhecido."""
    nome_norm = normalizar_nome(extrair_nome_completo_da_coluna(nome_coluna))
    if not nome_norm:
        return None, 0.0, "nome vazio"

    if nome_norm in nome_index:
        chaves = nome_index[nome_norm]
        if len(chaves) == 1:
            return chaves[0], 1.0, "exato"
        return None, 1.0, "nome repetido no efetivo"

    ranking = sorted(
        ((similaridade(nome_norm, k), k) for k in nome_index), reverse=True
    )
    if not ranking or ranking[0][0] < LIMIAR_SIMILARIDADE:
        melhor = ranking[0][0] if ranking else 0.0
        return None, melhor, "sem correspondência no efetivo"

    score, nome_ef = ranking[0]
    if len(ranking) > 1 and ranking[1][0] >= LIMIAR_SIMILARIDADE and score - ranking[1][0] < MARGEM_AMBIGUIDADE:
        return None, score, "ambíguo (mais de um nome parecido)"
    chaves = nome_index[nome_ef]
    if len(chaves) > 1:
        return None, score, "nome repetido no efetivo"
    return chaves[0], score, "aproximado"


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


def nome_completo_simples(dados: Dict) -> str:
    return f"{texto_simples(dados.get('posto_display', ''))} {texto_simples(dados.get('nome_display', ''))}".strip()


# =========================
# STATUS / PERÍODOS
# =========================
def classificar_status(resp: str) -> Tuple[str, int]:
    # Remove espaços e pontuação final (ex.: "Dispensa Luto." -> "Dispensa Luto")
    limpo = str(resp).strip().rstrip(".;:,").strip()
    rc = sem_acento_lower(limpo)
    # Opções padrão do formulário: correspondência EXATA
    if rc == "presente":                  return "Presente", 6
    if rc == "ausente":                   return "Ausente",  3
    if rc == "folga":                     return "Folga",    4
    if rc in DISPENSA_CHEFIA_OPCOES:      return "Dispensa pela Chefia", 5
    # Texto livre (campo "Outros"): preserva o texto digitado;
    # palavras-chave definem apenas a ordem no anúncio
    texto = limpo[:1].upper() + limpo[1:]
    return texto, ordem_status(texto)


def interpretar_resposta(valor: str) -> str:
    """
    Separa por vírgula apenas quando alguma parte é opção padrão do formulário
    (caixas de seleção). Texto livre com vírgula é mantido inteiro.
    """
    partes = [p.strip() for p in str(valor).split(",") if p.strip()]
    if len(partes) > 1 and any(
        sem_acento_lower(p).rstrip(".;:") in OPCOES_PADRAO or "ferias" in sem_acento_lower(p)
        for p in partes
    ):
        return min((classificar_status(p) for p in partes), key=lambda x: x[1])[0]
    return classificar_status(valor)[0]


def categoria_periodo(status: str) -> Optional[str]:
    """Tipo de afastamento que exige período, ou None."""
    s = sem_acento_lower(status)
    if s == "dispensa pela chefia":
        return None
    if "luto" in s or "nojo" in s:
        return "LUTO"
    if "ferias" in s:
        return "FÉRIAS"
    if "licenca" in s:
        return "LICENÇA"
    if "dispensa" in s:
        return "DISPENSA"
    return None


def precisa_periodo(status: str) -> bool:
    return categoria_periodo(status) is not None


def validar_periodo(inicio: date, fim: date) -> bool:
    return fim >= inicio


def formatar_periodo(inicio: date, fim: date) -> str:
    return f"{fmt_data(inicio)} a {fmt_data(fim)}"


def ordem_status(s: str) -> int:
    sl = sem_acento_lower(s)
    for kws, pri in STATUS_KEYWORDS:
        if any(k in sl for k in kws):
            return pri
    return 50


# =========================
# AFASTAMENTOS (aba da planilha)
# =========================
def preparar_afastamentos(valores: List[List[str]]) -> pd.DataFrame:
    colunas = ["linha", "chave", "matricula", "nome", "tipo", "descricao", "inicio", "fim"]
    if not valores or len(valores) < 2:
        return pd.DataFrame(columns=colunas)
    registros = []
    for i, r in enumerate(valores[1:], start=2):
        r = (list(r) + [""] * 6)[:6]
        ini, fim = parse_data_br(r[4]), parse_data_br(r[5])
        if not chave_matricula(r[0]) or ini is None or fim is None:
            continue
        registros.append({
            "linha": i, "chave": chave_matricula(r[0]), "matricula": r[0].strip(),
            "nome": r[1].strip(), "tipo": r[2].strip().upper(), "descricao": r[3].strip(),
            "inicio": ini, "fim": fim,
        })
    return pd.DataFrame(registros, columns=colunas)


def afastamentos_vigentes(df_af: pd.DataFrame, chave: str, dia: date, tipo: Optional[str] = None) -> pd.DataFrame:
    if df_af.empty:
        return df_af
    m = (df_af["chave"] == chave_matricula(chave)) & (df_af["inicio"] <= dia) & (df_af["fim"] >= dia)
    if tipo:
        m &= df_af["tipo"] == tipo
    return df_af[m].sort_values("linha", ascending=False)


def afastamento_existente(df_af: pd.DataFrame, chave: str, tipo: str, dia: date,
                          ini: date, fim: date) -> Optional[Dict]:
    """Registro do mesmo tipo vigente no dia ou que se sobrepõe ao novo período."""
    if df_af.empty:
        return None
    base = df_af[(df_af["chave"] == chave_matricula(chave)) & (df_af["tipo"] == tipo)]
    vig = base[(base["inicio"] <= dia) & (base["fim"] >= dia)]
    if vig.empty:
        vig = base[(base["inicio"] <= fim) & (base["fim"] >= ini)]
    if vig.empty:
        return None
    return vig.sort_values("linha", ascending=False).iloc[0].to_dict()


def valores_linha_afastamento(dados: Dict, tipo: str, status: str, ini: date, fim: date) -> List[str]:
    def seguro(s: str) -> str:
        s = str(s)
        return "'" + s if s[:1] in ("=", "+", "-", "@") else s
    return [
        dados["matricula"], seguro(nome_completo_simples(dados)), tipo,
        seguro(status), fmt_data(ini), fmt_data(fim),
    ]


def gravar_plano(plano: List[Dict]) -> Tuple[bool, str]:
    ops = []
    for p in plano:
        if p["acao"] == "incluir":
            ops.append({"op": "incluir", "valores": p["valores"]})
        elif p["acao"] == "atualizar":
            ops.append({"op": "atualizar", "linha": int(p["linha"]),
                        "matricula": p["matricula_linha"], "valores": p["valores"]})
    if not ops:
        return True, "Nenhuma alteração na aba AFASTAMENTOS."
    try:
        resp = chamar_api("gravar_afastamentos", ops=ops)
    except Exception as e:
        return False, f"Não foi possível gravar na aba AFASTAMENTOS: {e}"

    st.session_state.dados["afastamentos"] = preparar_afastamentos(resp.get("afastamentos", []))
    falhas = [r.get("erro") for r in resp.get("resultados", []) if not r.get("ok")]
    ok_qtd = len(ops) - len(falhas)
    msg = f"{ok_qtd} registro(s) gravado(s) na aba AFASTAMENTOS."
    if falhas:
        return False, msg + " Falhas: " + "; ".join(falhas)
    return True, msg


# =========================
# RANKING HIERÁRQUICO
# =========================
def limpar_para_ranking(texto: str) -> str:
    s = remover_asteriscos(str(texto)).upper().strip()
    s = remover_acentos(s)
    s = s.replace("º", "°")
    return re.sub(r"\s+", " ", s).strip()


def rank_hierarquico(dados: Dict) -> Tuple[int, int]:
    """(posto_rank, quadro_rank): reconvocados (QOR/QPR) após ativos do mesmo posto."""
    categoria  = dados.get("categoria", "")
    chave      = limpar_para_ranking(dados.get("posto_display", ""))
    chave      = re.sub(r"(\d+)°(TEN|SGT)", r"\1° \2", chave)
    quadro     = str(dados.get("quadro", "")).strip().upper()

    tabela = RANK_OFICIAIS if categoria == "OFICIAIS" else (
             RANK_PRACAS   if categoria == "PRAÇAS"   else {})

    if chave in tabela:
        posto_rank = tabela[chave]
    else:
        posto_rank = 999 if categoria == "CIVIS" else 900
        for k, v in tabela.items():
            if k in chave:
                posto_rank = v
                break

    return posto_rank, QUADRO_RANK.get(quadro, 0)


# =========================
# PROCESSAMENTO
# =========================
def processar_respostas(df_hoje: pd.DataFrame, nome_index: Dict, efetivo: Dict) -> Tuple[Dict, Dict]:
    """
    Vale o envio mais recente de cada seção (df_hoje deve vir ordenado do mais recente).
    Retorna (respostas {chave: {...}}, conferencia {...}).
    """
    colunas_militares = [c for c in df_hoje.columns if not eh_coluna_meta(c)]
    mapa = {c: encontrar_militar(c, nome_index) for c in colunas_militares}

    respostas = {}
    conf = {"aproximados": [], "nao_reconhecidos": [], "repetidos": []}
    aproximados_vistos = set()
    secoes_processadas = set()

    for _, row in df_hoje.iterrows():
        secao_form = str(row.get("Seção:", "")).strip()
        if secao_form in secoes_processadas:
            continue
        secoes_processadas.add(secao_form)

        for col in colunas_militares:
            valor = str(row[col]).strip()
            if not valor:
                continue
            chave, score, motivo = mapa[col]
            if chave is None:
                conf["nao_reconhecidos"].append({
                    "Coluna do formulário": col, "Seção (formulário)": secao_form,
                    "Resposta": valor, "Motivo": motivo,
                })
                continue
            if motivo == "aproximado" and col not in aproximados_vistos:
                aproximados_vistos.add(col)
                conf["aproximados"].append({
                    "Coluna do formulário": col,
                    "Associado a": nome_completo_simples(efetivo[chave]),
                    "Semelhança": f"{score:.0%}",
                })
            if chave in respostas:
                conf["repetidos"].append(
                    f"{nome_completo_simples(efetivo[chave])}: informado por "
                    f"'{respostas[chave]['secao_form']}' e por '{secao_form}' — "
                    f"vale o envio mais recente ('{respostas[chave]['secao_form']}')."
                )
                continue
            respostas[chave] = {
                "status": interpretar_resposta(valor),
                "dados": efetivo[chave],
                "secao_form": secao_form,
            }

    return respostas, conf


def divergencias_afastamentos(efetivo: Dict, respostas: Dict, df_af: pd.DataFrame, dia: date) -> List[str]:
    """Formulário prevalece; apenas alerta quando a aba AFASTAMENTOS indica outra situação."""
    alertas = []
    if df_af.empty:
        return alertas
    for chave, resp in respostas.items():
        vig = afastamentos_vigentes(df_af, chave, dia)
        if vig.empty:
            continue
        tipo_form = categoria_periodo(resp["status"])
        if tipo_form and (vig["tipo"] == tipo_form).any():
            continue
        r = vig.iloc[0]
        alertas.append(
            f"{nome_completo_simples(resp['dados'])}: aba AFASTAMENTOS indica {r['tipo']} "
            f"({formatar_periodo(r['inicio'], r['fim'])}), mas a seção informou "
            f"'{resp['status']}'. O anúncio segue o formulário."
        )
    return alertas


def organizar_categorias(efetivo: Dict, respostas: Dict, periodos: Dict) -> Tuple[Dict, List[Tuple[str, int]], List[str]]:
    categorias_dados = {
        cat: {"presentes": [], "afastamentos": {}, "total": 0}
        for cat in ["OFICIAIS", "PRAÇAS", "CIVIS"]
    }
    total_secao, faltam_secao = {}, {}

    for chave, dados in efetivo.items():
        categoria = dados["categoria"]
        categorias_dados[categoria]["total"] += 1
        secao = dados.get("secao") or "SEM SEÇÃO"
        total_secao[secao] = total_secao.get(secao, 0) + 1

        resposta = respostas.get(chave)
        if not resposta:
            faltam_secao.setdefault(secao, []).append(dados)
            continue

        status = str(resposta["status"]).strip()
        disp_base = formatar_nome_posto_somente_negritos(dados)
        rk = rank_hierarquico(dados)

        if precisa_periodo(status) and chave in periodos:
            ini, fim = periodos[chave]
            disp = f"{disp_base} - {formatar_periodo(ini, fim)}"
        else:
            disp = disp_base

        if "presente" in status.lower():
            categorias_dados[categoria]["presentes"].append((*rk, disp_base))
        else:
            categorias_dados[categoria]["afastamentos"].setdefault(status, []).append((*rk, disp))

    # Seção inteira sem resposta x militares isolados não informados
    secoes_sem_resposta, nao_informados = [], []
    for secao, lista in faltam_secao.items():
        if len(lista) == total_secao[secao]:
            secoes_sem_resposta.append((secao, len(lista)))
        else:
            nao_informados += [f"{formatar_nome_posto_somente_negritos(d)} ({secao})" for d in lista]

    return categorias_dados, secoes_sem_resposta, nao_informados


def gerar_anuncio(data_formatada: str, categorias_dados: Dict,
                  secoes_sem_resposta: List[Tuple[str, int]], nao_informados: List[str]) -> str:
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
            presentes = sorted(d["presentes"])
            partes.append(f"🔹{len(presentes)} Presentes:")
            partes += [f"    {i}. {t}" for i, (_, _, t) in enumerate(presentes, 1)]
            partes.append("")

        for status in sorted(d["afastamentos"], key=lambda s: (ordem_status(s), s)):
            lista = sorted(d["afastamentos"][status])
            partes.append(f"🔹{len(lista)} {status}")
            partes += [f"    {i}. {t}" for i, (_, _, t) in enumerate(lista, 1)]
            partes.append("")

        partes.append("")

    if secoes_sem_resposta:
        itens = sorted(secoes_sem_resposta, key=lambda x: (-x[1], x[0]))
        partes.append(f"❌ Seções sem resposta ({len(itens)}):")
        for secao, qtd in itens:
            partes.append(f"➡️ {secao} ({qtd} servidores no total);")
        partes.append("")

    if nao_informados:
        partes.append(f"⚠️ Servidores não informados ({len(nao_informados)}):")
        partes += [f"➡️ {t};" for t in sorted(nao_informados)]
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
    return "\n".join(partes)


# =========================
# UI — PERÍODOS
# =========================
def etapa_periodos(afastados: List[Tuple], df_af: pd.DataFrame, hoje: date) -> bool:
    """Conduz a etapa de períodos. Retorna True quando o anúncio pode ser exibido."""
    etapa = st.session_state.etapa_periodos

    if etapa == "editar":
        st.write("Confira os períodos e clique em **Aplicar períodos**. "
                 "Os já registrados na aba AFASTAMENTOS vêm preenchidos.")
        with st.form("form_periodos"):
            novos, erros = {}, []
            for chave, dados, status, tipo in afastados:
                st.markdown(f"**{formatar_nome_posto_somente_negritos(dados)}** — _{status}_ ({tipo})")
                padrao = st.session_state.periodos_inseridos.get(chave)
                if padrao is None:
                    vig = afastamentos_vigentes(df_af, chave, hoje, tipo)
                    if not vig.empty:
                        padrao = (vig.iloc[0]["inicio"], vig.iloc[0]["fim"])
                        st.caption(f"📌 Registrado na aba: {formatar_periodo(*padrao)}")
                    else:
                        padrao = (hoje, hoje)
                        st.caption("🆕 Sem registro vigente na aba — informe o período.")
                c1, c2 = st.columns(2)
                ini = c1.date_input("Início", value=padrao[0], key=f"ini_{chave}", format="DD/MM/YYYY")
                fim = c2.date_input("Fim", value=padrao[1], key=f"fim_{chave}", format="DD/MM/YYYY")
                if not validar_periodo(ini, fim):
                    erros.append(f"{nome_completo_simples(dados)}: fim anterior ao início.")
                novos[chave] = (ini, fim)
                st.markdown("---")
            enviado = st.form_submit_button("✅ Aplicar períodos")

        if enviado:
            if erros:
                for e in erros:
                    st.error(e)
                return False
            plano = []
            for chave, dados, status, tipo in afastados:
                ini, fim = novos[chave]
                valores = valores_linha_afastamento(dados, tipo, status, ini, fim)
                if chave.startswith("SEM-MATRICULA"):
                    plano.append({"acao": "ignorar", "nome": nome_completo_simples(dados),
                                  "motivo": "sem matrícula no efetivo"})
                    continue
                ex = afastamento_existente(df_af, chave, tipo, hoje, ini, fim)
                if ex is None:
                    plano.append({"acao": "incluir", "valores": valores})
                elif (ex["inicio"], ex["fim"]) == (ini, fim):
                    continue
                else:
                    plano.append({
                        "acao": "decidir", "chave": chave, "nome": nome_completo_simples(dados),
                        "tipo": tipo, "linha": int(ex["linha"]), "matricula_linha": ex["matricula"],
                        "antigo": formatar_periodo(ex["inicio"], ex["fim"]),
                        "novo": formatar_periodo(ini, fim), "valores": valores,
                    })
            st.session_state.periodos_inseridos = novos
            st.session_state.plano_gravacao = plano
            if any(p["acao"] == "decidir" for p in plano):
                st.session_state.etapa_periodos = "confirmar"
            else:
                ok, msg = gravar_plano(plano)
                st.session_state.msg_gravacao = (ok, msg, plano)
                st.session_state.etapa_periodos = "aplicado"
            st.rerun()
        return False

    if etapa == "confirmar":
        st.warning("Alguns períodos diferem do que já está registrado na aba AFASTAMENTOS. "
                   "Escolha, em cada caso, o que fazer:")
        escolhas = {}
        for i, p in enumerate(st.session_state.plano_gravacao):
            if p["acao"] != "decidir":
                continue
            st.markdown(f"**{p['nome']}** ({p['tipo']}) — na aba: {p['antigo']} → novo: {p['novo']}")
            escolhas[i] = st.radio(
                "O que fazer?", ["Atualizar a linha existente", "Criar nova linha"],
                key=f"decisao_{i}", horizontal=True,
            )
        c1, c2 = st.columns(2)
        if c1.button("💾 Confirmar e gravar", type="primary"):
            plano = []
            for i, p in enumerate(st.session_state.plano_gravacao):
                if p["acao"] == "decidir":
                    p = dict(p, acao="atualizar" if escolhas[i].startswith("Atualizar") else "incluir")
                plano.append(p)
            ok, msg = gravar_plano(plano)
            st.session_state.msg_gravacao = (ok, msg, plano)
            st.session_state.etapa_periodos = "aplicado"
            st.rerun()
        if c2.button("↩️ Voltar e editar"):
            st.session_state.etapa_periodos = "editar"
            st.rerun()
        return False

    # aplicado
    if st.session_state.msg_gravacao:
        ok, msg, plano = st.session_state.msg_gravacao
        (st.success if ok else st.error)(("✅ " if ok else "❌ ") + msg)
        for p in plano:
            if p["acao"] == "ignorar":
                st.warning(f"{p['nome']}: período não gravado ({p['motivo']}).")
        if not ok:
            st.caption("O anúncio usa os períodos informados mesmo assim.")
    if st.button("✏️ Editar períodos"):
        st.session_state.etapa_periodos = "editar"
        st.session_state.msg_gravacao = None
        st.rerun()
    return True


# =========================
# UI PRINCIPAL
# =========================
def main():
    st.set_page_config(page_title="Anúncio CSC-PM", page_icon="📢")
    init_session_state()

    hoje = agora().date()
    data_formatada = fmt_data(hoje)

    st.title(f"GERADOR DE ANÚNCIO DE PRESENÇA CSC-PM {VERSAO}")
    st.markdown("---")

    with st.sidebar:
        st.subheader("⚙️ Controles")
        if st.button("🔄 Reset completo"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()
        st.markdown("---")
        st.markdown(
            "**📋 Abas usadas na planilha:**\n"
            "- `Respostas ao formulário 1`\n"
            "- `EFETIVO CSC` — SEÇÃO | NÚMERO | P / G | QUADRO | NOME\n"
            "- `AFASTAMENTOS` — criada e mantida pelo app\n\n"
            "Use `*TEXTO*` no efetivo para negrito no WhatsApp."
        )
        st.caption(f"{VERSAO} — leitura e gravação via Apps Script")

    # ── 1) Dados ──────────────────────────────────────────────
    st.subheader("1️⃣ Dados da planilha")
    recarregar = st.button("🔄 Atualizar dados da planilha")
    if st.session_state.dados is None or recarregar or st.session_state.get("data_carga") != hoje:
        try:
            with st.spinner("Lendo a planilha..."):
                st.session_state.dados = carregar_dados(hoje)
            st.session_state.carregado_em = agora()
            st.session_state.data_carga = hoje
            resetar_periodos()
        except Exception as e:
            st.error(f"❌ Erro ao ler a planilha: {e}")
            st.stop()
    st.caption(f"Dados lidos às {st.session_state.carregado_em.strftime('%H:%M:%S')} de {data_formatada}.")
    dados = st.session_state.dados

    # ── 2) Efetivo ────────────────────────────────────────────
    st.markdown("---")
    st.subheader("2️⃣ Efetivo CSC")
    try:
        efetivo, nome_index, alertas_efetivo = carregar_efetivo_do_df(dados["efetivo"])
    except Exception as e:
        st.error(f"❌ Erro na aba de efetivo: {e}")
        st.stop()
    cont = {c: sum(1 for d in efetivo.values() if d["categoria"] == c) for c in ["OFICIAIS", "PRAÇAS", "CIVIS"]}
    st.success(f"✅ {len(efetivo)} servidores ({cont['OFICIAIS']} oficiais | "
               f"{cont['PRAÇAS']} praças | {cont['CIVIS']} civis)")
    for a in alertas_efetivo:
        st.warning(a)

    # ── 3) Respostas ──────────────────────────────────────────
    st.markdown("---")
    st.subheader("3️⃣ Respostas do dia")
    df_form = dados["formulario"]
    faltando = {"Carimbo de data/hora", "Data do anúncio", "Seção:"} - set(df_form.columns)
    if faltando:
        st.error(f"❌ Colunas ausentes na aba de respostas: {', '.join(sorted(faltando))}")
        st.stop()

    df_form = df_form.copy()
    df_form["_carimbo"] = df_form["Carimbo de data/hora"].map(parse_datahora_br)
    df_form["_data"] = df_form["Data do anúncio"].map(parse_data_br)
    df_hoje = df_form[df_form["_data"] == hoje].drop(columns=["_data"])
    if df_hoje.empty:
        st.warning(f"⚠️ Nenhuma resposta para {data_formatada}.")
        st.stop()
    df_hoje = df_hoje.sort_values("_carimbo", ascending=False).drop(columns=["_carimbo"])
    st.success(f"✅ {len(df_hoje)} envio(s) para {data_formatada} "
               f"({df_hoje['Seção:'].nunique()} seção(ões)).")

    respostas, conf = processar_respostas(df_hoje, nome_index, efetivo)
    df_af = dados["afastamentos"]
    divergencias = divergencias_afastamentos(efetivo, respostas, df_af, hoje)

    qtd_avisos = len(conf["nao_reconhecidos"]) + len(conf["repetidos"]) + len(divergencias)
    titulo = "🔎 Conferência"
    titulo += f" — {qtd_avisos} ponto(s) de atenção" if qtd_avisos else " — sem pendências"
    with st.expander(titulo, expanded=qtd_avisos > 0):
        if conf["nao_reconhecidos"]:
            st.markdown("**❗ Respostas não associadas a ninguém do efetivo** (não entram no anúncio):")
            st.dataframe(pd.DataFrame(conf["nao_reconhecidos"]), hide_index=True, use_container_width=True)
        if divergencias:
            st.markdown("**⚠️ Divergências com a aba AFASTAMENTOS:**")
            for d in divergencias:
                st.write(f"• {d}")
        if conf["repetidos"]:
            st.markdown("**⚠️ Militar informado por mais de uma seção:**")
            for d in conf["repetidos"]:
                st.write(f"• {d}")
        if conf["aproximados"]:
            st.markdown("**ℹ️ Associações por semelhança de nome** (confira):")
            st.dataframe(pd.DataFrame(conf["aproximados"]), hide_index=True, use_container_width=True)
        if not qtd_avisos and not conf["aproximados"]:
            st.write("Todas as respostas foram associadas exatamente ao efetivo.")

    # ── 4) Períodos ───────────────────────────────────────────
    afastados = [
        (chave, r["dados"], r["status"], categoria_periodo(r["status"]))
        for chave, r in respostas.items() if precisa_periodo(r["status"])
    ]
    afastados.sort(key=lambda x: (rank_hierarquico(x[1]), x[1]["nome_norm"]))
    assinatura = tuple((a[0], a[2]) for a in afastados)
    if st.session_state.get("assinatura_afastados") != assinatura:
        st.session_state.assinatura_afastados = assinatura
        resetar_periodos()

    st.markdown("---")
    st.subheader("4️⃣ Períodos de afastamento")
    if afastados:
        liberado = etapa_periodos(afastados, df_af, hoje)
    else:
        st.info("Nenhum servidor com férias, licença, luto ou dispensa hoje.")
        liberado = True

    if not liberado:
        st.info("📢 O anúncio será exibido após aplicar os períodos.")
        st.stop()

    # ── 5) Anúncio ────────────────────────────────────────────
    periodos = st.session_state.periodos_inseridos if afastados else {}
    categorias_dados, secoes_sem_resposta, nao_informados = organizar_categorias(efetivo, respostas, periodos)
    anuncio = gerar_anuncio(data_formatada, categorias_dados, secoes_sem_resposta, nao_informados)

    st.markdown("---")
    st.subheader("📢 ANÚNCIO GERADO:")
    st.code(anuncio, language="text")

    st.download_button(
        label="💾 Baixar Anúncio",
        data=anuncio.encode("utf-8"),
        file_name=f"anuncio_presenca_{hoje.strftime('%Y%m%d')}.txt",
        mime="text/plain",
    )


if __name__ == "__main__":
    main()
