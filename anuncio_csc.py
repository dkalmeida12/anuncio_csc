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

# Regex compilado para melhor performance
SHEET_ID_PATTERN = re.compile(r"/spreadsheets/d/([a-zA-Z0-9-_]+)")
POSTO_PATTERNS = [
    (re.compile(r'^[\s]*ASPM[\s]+', re.IGNORECASE), ''),
    (re.compile(r'^[\s]*\d+[º°][\s]*', re.IGNORECASE), ''),
    (re.compile(r'^[\s]*(TEN[\s]*CEL|MAJ|CAP|SUB[\s]*TENENTE|SUBTENENTE|TEN|SGT|CB)[\s]+', re.IGNORECASE), ''),
    (re.compile(r'^[\s]*\d+[º°]?(TEN|SGT)[\s]+', re.IGNORECASE), '')
]

def extrair_sheet_id(url: str) -> str:
    """Extrai ID da planilha Google Sheets da URL."""
    m = SHEET_ID_PATTERN.search(str(url))
    return m.group(1) if m else ""

def baixar_sheets_publico_xlsx(sheet_url: str) -> bytes:
    """Baixa planilha pública do Google Sheets em formato XLSX."""
    sheet_id = extrair_sheet_id(sheet_url)
    if not sheet_id:
        raise ValueError("Não foi possível extrair o SHEET_ID da URL informada.")
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    r = requests.get(export_url, timeout=30)
    r.raise_for_status()
    return r.content

# =========================
# EFETIVO CSC-PM (INTEGRADO NO CÓDIGO)
# =========================
EFETIVO_CSC = """SEÇÃO,NÚMERO,P  / G,QUADRO,NOME
CHEFE,126.554-5,*TEN CEL*,QOPM,*LEONARDO* de *CASTRO* Ferreira
SUBCHEFE,089.655-5,*MAJ*,QOR,Jorge Aparecido *GOMES*
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

# Mapeamento de quadros para categorias (constante)
QUADRO_CATEGORIA = {
    'QOPM': 'OFICIAIS', 'QOR': 'OFICIAIS', 'QOC': 'OFICIAIS',
    'QPR': 'PRAÇAS', 'QPPM': 'PRAÇAS', 'QPE': 'PRAÇAS',
    'CIVIL': 'CIVIS'
}

# Palavras-chave para classificação de status (ordenadas por prioridade)
STATUS_KEYWORDS = [
    (['férias', 'ferias'], 1),
    (['licença', 'licenca'], 2),
    (['ausente'], 3),
    (['folga'], 4),
    (['dispensa'], 5),
    (['presente'], 6)
]

# =========================
# Session State Init
# =========================
def init_session_state():
    """Inicializa variáveis do session state."""
    defaults = {
        "df_formulario": None,
        "fonte_ok": False,
        "periodos_aplicados": False,
        "periodos_inseridos": {},
        "periodos_memoria": {}
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

# =========================
# Funções auxiliares otimizadas
# =========================
def remover_asteriscos(s: str) -> str:
    """Remove asteriscos de uma string."""
    return s.replace('*', '') if s else ""

def remover_acentos(s: str) -> str:
    """Remove acentos de uma string."""
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))

@st.cache_data
def normalizar_nome(nome: str) -> str:
    """
    Normalização para matching (cached para performance).
    Remove *, converte para upper, remove acentos e pontuação.
    """
    if pd.isna(nome):
        return ""
    s = remover_asteriscos(str(nome).strip().upper())
    s = remover_acentos(s)
    s = re.sub(r"[^A-Z\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def normalizar_posto(posto: str) -> str:
    """Normaliza posto/graduação mantendo asteriscos para formatação WhatsApp."""
    return str(posto).strip().replace('º', '°').replace('1º', '1°').replace('2º', '2°').replace('3º', '3°')

def extrair_nome_completo_da_coluna(nome_coluna: str) -> str:
    """Extrai nome completo do militar do cabeçalho da coluna."""
    s = str(nome_coluna).strip()
    
    # Caso especial: " PM "
    idx = s.upper().rfind(" PM ")
    if idx != -1:
        return s[idx + 4:].strip()
    
    # Aplicar padrões regex pré-compilados
    for pattern, repl in POSTO_PATTERNS:
        s = pattern.sub(repl, s)
    
    return s.strip()

def similaridade(a: str, b: str) -> float:
    """Calcula similaridade entre duas strings."""
    return SequenceMatcher(None, a, b).ratio()

def encontrar_militar(nome_extraido: str, efetivo_dict: Dict, limiar: float = 0.88) -> Tuple[Optional[str], Optional[Dict]]:
    """Encontra militar no dicionário de efetivo usando matching exato ou por similaridade."""
    nome_norm = normalizar_nome(nome_extraido)
    
    # Busca exata (mais rápida)
    if nome_norm in efetivo_dict:
        return nome_norm, efetivo_dict[nome_norm]
    
    # Busca por similaridade
    melhor_key, melhor_score = max(
        ((key, similaridade(nome_norm, key)) for key in efetivo_dict),
        key=lambda x: x[1],
        default=(None, 0.0)
    )
    
    return (melhor_key, efetivo_dict[melhor_key]) if melhor_score >= limiar else (None, None)

def classificar_status(resp: str) -> Tuple[str, int]:
    """Classifica status da resposta e retorna prioridade."""
    resp_lower = resp.lower()
    
    # Mapeamento direto para respostas comuns
    if resp_lower == 'presente':
        return "Presente", 6
    if resp_lower == 'ausente':
        return "Ausente", 3
    if resp_lower == 'folga':
        return "Folga", 4
    if 'dispensa' in resp_lower:
        return "Dispensa pela Chefia", 5
    
    # Busca por palavras-chave
    for keywords, priority in STATUS_KEYWORDS:
        if any(kw in resp_lower for kw in keywords):
            return resp, priority
    
    return resp, 50

def precisa_periodo(status: str) -> bool:
    """Verifica se o status requer informação de período."""
    sl = status.lower()
    return 'férias' in sl or 'ferias' in sl or 'licença' in sl or 'licenca' in sl

def validar_periodo(inicio: date, fim: date) -> bool:
    """Valida se o período é consistente."""
    return fim >= inicio

def formatar_periodo(inicio: date, fim: date) -> str:
    """Formata período para exibição."""
    return f"{inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}"

def ordem_status(s: str) -> int:
    """Retorna ordem de prioridade para ordenação de status."""
    sl = s.lower()
    for keywords, priority in STATUS_KEYWORDS:
        if any(kw in sl for kw in keywords):
            return priority
    return 50

# =========================
# Processamento de dados
# =========================
@st.cache_data
def carregar_efetivo() -> Dict:
    """Carrega e processa dados do efetivo (cached)."""
    df_efetivo = pd.read_csv(io.StringIO(EFETIVO_CSC))
    
    # Strip em todas as colunas
    colunas = ['SEÇÃO', 'NÚMERO', 'P  / G', 'QUADRO', 'NOME']
    for col in colunas:
        df_efetivo[col] = df_efetivo[col].astype(str).str.strip()
    
    efetivo_dict = {}
    
    for _, row in df_efetivo.iterrows():
        quadro = row['QUADRO'].upper()
        categoria = QUADRO_CATEGORIA.get(quadro)
        
        if not categoria:
            continue
        
        nome_display = row['NOME']
        posto_display = normalizar_posto(row['P  / G'])
        nome_norm = normalizar_nome(nome_display)
        
        efetivo_dict[nome_norm] = {
            'categoria': categoria,
            'posto_display': posto_display,
            'nome_display': nome_display,
            'quadro': quadro,
            'secao': row['SEÇÃO'].upper()
        }
    
    return efetivo_dict

def processar_respostas(df_hoje: pd.DataFrame, efetivo_dict: Dict) -> Dict:
    """Processa respostas do formulário e retorna dicionário de status."""
    respostas_dict = {}
    secoes_processadas = set()
    colunas_militares = df_hoje.columns[4:]
    
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
            chave_efetivo, militar_encontrado = encontrar_militar(nome_militar, efetivo_dict)
            
            if not militar_encontrado:
                continue
            
            # Processar múltiplas respostas separadas por vírgula
            respostas = [r.strip() for r in str(valor).strip().split(',') if r.strip()]
            candidatos = [classificar_status(resp) for resp in respostas]
            
            if candidatos:
                # Seleciona resposta com maior prioridade (menor número)
                status_texto_exato = min(candidatos, key=lambda x: x[1])[0]
                respostas_dict[chave_efetivo] = {
                    'status': status_texto_exato,
                    'dados': militar_encontrado
                }
    
    return respostas_dict

def organizar_categorias(efetivo_dict: Dict, respostas_dict: Dict, periodos_inseridos: Dict) -> Tuple[Dict, Dict, List[str]]:
    """Organiza dados por categoria e identifica faltantes."""
    categorias_dados = {
        cat: {'presentes': [], 'afastamentos': {}, 'total': 0}
        for cat in ['OFICIAIS', 'PRAÇAS', 'CIVIS']
    }
    
    faltantes_por_secao = {}
    militares_nao_informados = []
    
    for nome_norm, dados in efetivo_dict.items():
        categoria = dados['categoria']
        categorias_dados[categoria]['total'] += 1
        
        resposta = respostas_dict.get(nome_norm)
        
        if not resposta:
            secao = dados.get('secao', 'SEM SEÇÃO')
            faltantes_por_secao[secao] = faltantes_por_secao.get(secao, 0) + 1
            militares_nao_informados.append(f"{dados['posto_display']} {dados['nome_display']} ({secao})")
            continue
        
        status = str(resposta['status']).strip()
        posto_nome_display = f"{dados['posto_display']} {dados['nome_display']}"
        
        # Adicionar período se aplicável
        if precisa_periodo(status) and nome_norm in periodos_inseridos:
            ini, fim = periodos_inseridos[nome_norm]
            posto_nome_saida = f"{posto_nome_display} - {formatar_periodo(ini, fim)}"
        else:
            posto_nome_saida = posto_nome_display
        
        # Classificar como presente ou afastamento
        if 'presente' in status.lower() or status == "Presente":
            categorias_dados[categoria]['presentes'].append(posto_nome_display)
        else:
            categorias_dados[categoria]['afastamentos'].setdefault(status, []).append(posto_nome_saida)
    
    return categorias_dados, faltantes_por_secao, militares_nao_informados

def gerar_anuncio(data_formatada: str, categorias_dados: Dict, faltantes_por_secao: Dict) -> Tuple[str, int, int]:
    """Gera texto do anúncio de presença."""
    anuncio_parts = [
        "Bom dia!",
        "Segue anúncio do dia",
        "",
        "Anúncio CSC-PM",
        data_formatada,
        ""
    ]
    
    total_militares = 0
    total_civis = 0
    
    for categoria in ['OFICIAIS', 'PRAÇAS', 'CIVIS']:
        dados_cat = categorias_dados[categoria]
        
        if categoria == 'CIVIS':
            total_civis = len(dados_cat['presentes'])
        else:
            total_militares += len(dados_cat['presentes'])
        
        anuncio_parts.extend([
            f"*{categoria}*",
            "Efetivo total: ",
            f"🔸{dados_cat['total']} - CSC-PM",
                ])
        
        if dados_cat['presentes']:
            anuncio_parts.append(f"🔹{len(dados_cat['presentes'])} Presentes:")
            anuncio_parts.extend(f"    {i}. {nome}" for i, nome in enumerate(dados_cat['presentes'], 1))
            anuncio_parts.append("")
        
        for status in sorted(dados_cat['afastamentos'].keys(), key=ordem_status):
            lista = dados_cat['afastamentos'][status]
            anuncio_parts.append(f"🔹{len(lista)} {status}")
            anuncio_parts.extend(f"    {i}. {info}" for i, info in enumerate(lista, 1))
            anuncio_parts.append("")
        
        anuncio_parts.append("")
    
    if faltantes_por_secao:
        itens = sorted(faltantes_por_secao.items(), key=lambda x: (-x[1], x[0]))
        anuncio_parts.append(f"❌ Seções sem resposta ({len(itens)}):")
        anuncio_parts.extend(f"➡️ {secao}({qtd} servidores no total);" for secao, qtd in itens)
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
# UI Principal
# =========================
def main():
    init_session_state()
    
    st.title("GERADOR DE ANÚNCIO DE PRESENÇA CSC-PM v3.9")
    st.markdown("---")
    
    # Sidebar com controles
    with st.sidebar:
        if st.button("Limpar carregamento (reset)"):
            for key in ["df_formulario", "fonte_ok", "periodos_aplicados", "periodos_inseridos"]:
                st.session_state[key] = None if key == "df_formulario" else (False if "ok" in key or "aplicados" in key else {})
            st.rerun()
        
        if st.button("Limpar memória de períodos"):
            st.session_state.periodos_memoria = {}
            st.success("Memória de períodos limpa.")
            st.rerun()
    
    # Seção 1: Carregar planilha
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
    
    # Processamento principal
    data_atual = datetime.now()
    data_formatada = data_atual.strftime("%d/%m/%Y")
    
    efetivo_dict = carregar_efetivo()
    
    # Validar colunas obrigatórias
    colunas_obrigatorias = {'Carimbo de data/hora', 'Data do anúncio', 'Seção:'}
    faltando = colunas_obrigatorias - set(df_formulario.columns.astype(str))
    if faltando:
        st.error(f"❌ A planilha não possui as colunas obrigatórias: {', '.join(sorted(faltando))}")
        st.stop()
    
    # Filtrar registros de hoje
    df_formulario['Carimbo de data/hora'] = pd.to_datetime(df_formulario['Carimbo de data/hora'])
    df_formulario['Data do anúncio'] = pd.to_datetime(df_formulario['Data do anúncio'])
    df_hoje = df_formulario[df_formulario['Data do anúncio'].dt.date == data_atual.date()].copy()
    
    st.markdown("---")
    st.subheader("2) Leitura das respostas")
    
    if df_hoje.empty:
        st.warning(f"⚠️ ATENÇÃO: Não há registros para a data {data_formatada}")
        st.info("Verifique se a 'Data do anúncio' no formulário corresponde à data de hoje.")
        st.stop()
    
    st.success(f"✅ Encontrados {len(df_hoje)} registro(s) para {data_formatada}")
    df_hoje = df_hoje.sort_values('Carimbo de data/hora', ascending=False)
    
    respostas_dict = processar_respostas(df_hoje, efetivo_dict)
    
    # Seção 3: Períodos
    afastados = [
        (chave, resp['dados'], resp['status'])
        for chave, resp in respostas_dict.items()
        if precisa_periodo(resp['status'])
    ]
    
    st.markdown("---")
    st.subheader("3) Informar períodos (Férias / Licença)")
    
    if afastados and not st.session_state.periodos_aplicados:
        st.write("Preencha início e fim e clique em **Aplicar períodos**. No anúncio será exibido: `POSTO NOME - dd/mm/aaaa a dd/mm/aaaa`")
        
        with st.form("form_periodos"):
            novos_periodos = {}
            erros = []
            
            for chave_norm, dados, status in afastados:
                posto_nome_display = f"{dados['posto_display']} {dados['nome_display']}"
                st.markdown(f"**{posto_nome_display}**  \n_{status}_")
                
                # Pré-preenchimento por memória
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
            
            submitted = st.form_submit_button("Aplicar períodos")
        
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
    
    # Organizar dados e gerar anúncio
    categorias_dados, faltantes_por_secao, militares_nao_informados = organizar_categorias(
        efetivo_dict, respostas_dict, periodos_inseridos
    )
    
    anuncio, total_militares, total_civis = gerar_anuncio(data_formatada, categorias_dados, faltantes_por_secao)
    
    # Exibir resultados
    st.markdown("---")
    st.subheader("📢 ANÚNCIO GERADO:")
    st.code(anuncio, language='text')
    
    if faltantes_por_secao:
        with st.expander("Ver militares que não responderam (conferência)"):
            for item in sorted(militares_nao_informados):
                st.write(f"• {item}")
    
    st.download_button(
        label="Baixar Anúncio de Presença",
        data=anuncio.encode('utf-8'),
        file_name="anuncio_presenca.txt",
        mime="text/plain"
    )
    
    st.success("✅ PROCESSO CONCLUÍDO!")

if __name__ == "__main__":
    main()
