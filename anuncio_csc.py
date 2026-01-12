import pandas as pd
from datetime import datetime
import re
import unicodedata
from difflib import SequenceMatcher
import streamlit as st

# EFETIVO CSC-PM (INTEGRADO NO CÓDIGO)
EFETIVO_CSC = """SEÇÃO,NÚMERO,P  / G,QUADRO,NOME
CHEFE,126.554-5,TEN CEL,QOPM,LEONARDO DE CASTRO FERREIRA
SUBCHEFE,089.655-5,MAJ,QOR,JORGE APARECIDO GOMES
LICITAÇÃO,161.300-9,CAP,QOPM,THIAGO FERNANDES PALMEIRA
LICITAÇÃO,100.433-2,2ºTEN,QOR,CLAUDIO MARCIO DA SILVA
LICITAÇÃO,087.650-8,SUBTEN,QPR,SÉRGIO BERNARDINO DE SENA
LICITAÇÃO,154.178-8,2ºSGT,QPPM,THIAGO LUIZ TEIXEIRA
COMPRAS,134.166-8,CAP,QOPM,SAMUEL LUIZ VIEIRA
COMPRAS,135.147-7,2ºTEN,QOC,CLEUBER FERREIRA DA SILVA
COMPRAS,147.720-7,3º SGT,QPE,HERBERT DIOGO FRADE GARBAZZA
P1,166.850-8,1º TEN,QOPM,DIEGO KUKIYAMA DE ALMEIDA
P1,087.768-8,1ºSGT,QPR,GLAUCO ALMEIDA BRAZ
P1,094.907-3,2ºSGT,QPR,ALEXANDRE AUGUSTO CORREA
P1,140.204-9,3ºSGT,QPPM,LEONARDO GOMES DA COSTA
P1,144.105-4,3ºSGT,QPPM,MAURO JACOB DE GOUVEIA QUIRINO
P1,181.220-5,3ºSGT,QPPM,NÚBIA APARECIDA RIBEIRO
P1,174.777-3,CB,QPPM,ANA CLÁUDIA TAVARES CAETANO
P1,167.318-5,ASPM,CIVIL,MARA CRISTINA DUARTE PEREIRA
SOFI,149.668-6,CAP,QOPM,DIOGO DA SILVA ROSA
SOFI,134.606-3,1ºTEN,QOC,VALTER ADRIANO DOS SANTOS
SOFI,134.927-3,3º SGT,QPPM,WALITON KELITON DA CRUZ
SOFI,146.417-1,3º SGT,QPPM,TIAGO HENRIQUE DA SILVA
SOFI,146.299-3,3º SGT,QPPM,GUSTAVO GUIMARÃES AFEITO
ALMOX,099.519-1,2ºTEN,QOR,WALMIR MÁRCIO DA CRUZ
ALMOX,099.309-7,1ºSGT,QPR,OMAIR CELSO DOS SANTOS
ALMOX,113.505-2,1ºSGT,QPR,CARLOS LAÉRCIO DE SOUZA
ALMOX,167.118-9,ASPM,CIVIL,DANIELLE GOMES FIGUEIROA
S PRODUÇÃO GRÁFICA,094.227-6,2ºTEN,QOR,VILMO GONÇALVES LEMOS
S MANUTENÇÃO,087.957-7,2ºTEN,QOR,JOAQUIM ARAÚJO DE OLIVEIRA
S MANUTENÇÃO,102.773-9,2ºSGT,QPR,NIVAL NEVES DE CARVALHO
S MANUTENÇÃO,090.803-8,2ºSGT,QPR,ARNALDO BENTO PEREIRA
S MANUTENÇÃO,097.538-3,3ºSGT,QPR,CARLOS R SANTIAGO DOS SANTOS
S MANUTENÇÃO,127.860-5,3º SGT,QPPM,WAGNER VITOR DOS SANTOS"""

st.title("GERADOR DE ANÚNCIO DE PRESENÇA CSC-PM v3.2")
st.markdown("---")

# Carregar efetivo
df_efetivo = pd.read_csv(pd.io.common.StringIO(EFETIVO_CSC))

# Upload do formulário
st.write("📋 FAÇA O UPLOAD DA PLANILHA DO GOOGLE FORMULÁRIOS (XLS/XLSX):")
uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xls", "xlsx"])

if uploaded_file is not None:
    st.write("📊 Processando dados...")
    df_formulario = pd.read_excel(uploaded_file)

    data_atual = datetime.now()
    data_formatada = data_atual.strftime("%d/%m/%Y")

    # ----------------------------
    # Normalizações / Matching
    # ----------------------------
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

    # ----------------------------
    # Período (opcional)
    # ----------------------------
    def extrair_periodo(texto):
        if pd.isna(texto):
            return None
        texto = str(texto)

        padrao1 = r'(\d{2}[a-zA-Z]{3})[\s]*[àaáÀÁ][\s]*(\d{2}[a-zA-Z]{3})'
        padrao2 = r'(\d{2}/\d{2})[\s]*[àaáÀÁ][\s]*(\d{2}/\d{2})'

        match = re.search(padrao1, texto, re.IGNORECASE)
        if match:
            return f"{match.group(1)} à {match.group(2)}"

        match = re.search(padrao2, texto)
        if match:
            return f"{match.group(1)} à {match.group(2)}"

        return None

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

    # ----------------------------
    # Processar efetivo -> dict
    # ----------------------------
    efetivo_dict = {}
    for _, row in df_efetivo.iterrows():
        nome_completo = str(row['NOME']).strip()
        quadro = str(row['QUADRO']).strip().upper()
        posto_grad = normalizar_posto(str(row['P  / G']))

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
                'quadro': quadro
            }

    # ----------------------------
    # Processar formulário (dia atual)
    # ----------------------------
    df_formulario['Carimbo de data/hora'] = pd.to_datetime(df_formulario['Carimbo de data/hora'])
    df_formulario['Data do anúncio'] = pd.to_datetime(df_formulario['Data do anúncio'])

    df_hoje = df_formulario[df_formulario['Data do anúncio'].dt.date == data_atual.date()].copy()

    if df_hoje.empty:
        st.warning(f"⚠️ ATENÇÃO: Não há registros para a data {data_formatada}")
        st.info("Verifique se a 'Data do anúncio' no formulário corresponde à data de hoje.")
    else:
        st.success(f"✅ Encontrados {len(df_hoje)} registro(s) para {data_formatada}")

    df_hoje = df_hoje.sort_values('Carimbo de data/hora', ascending=False)

    # ----------------------------
    # Coletar respostas
    # ----------------------------
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

            nome_coluna = str(col).strip()
            nome_militar = extrair_nome_completo_da_coluna(nome_coluna)

            chave_efetivo, militar_encontrado = encontrar_militar(nome_militar, efetivo_dict, limiar=0.88)
            if not militar_encontrado:
                continue

            valor_str = str(valor).strip()
            respostas = [r.strip() for r in valor_str.split(',') if r.strip()]

            candidatos = []
            melhor_periodo = None

            for resp in respostas:
                resp_lower = resp.lower()
                periodo = extrair_periodo(resp)

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

                if periodo and not melhor_periodo:
                    melhor_periodo = periodo

            if not candidatos:
                continue

            candidatos.sort(key=lambda x: x[1])
            status_texto_exato = candidatos[0][0]

            respostas_dict[chave_efetivo] = {
                'status': status_texto_exato,
                'campo_outro': melhor_periodo,
                'dados': militar_encontrado
            }

    # ----------------------------
    # Organizar por categoria / status dinâmico
    # ----------------------------
    categorias_dados = {
        'OFICIAIS': {'presentes': [], 'afastamentos': {}, 'total': 0},
        'PRAÇAS': {'presentes': [], 'afastamentos': {}, 'total': 0},
        'CIVIS': {'presentes': [], 'afastamentos': {}, 'total': 0}
    }

    militares_nao_informados = []

    for nome_norm, dados in efetivo_dict.items():
        categoria = dados['categoria']
        categorias_dados[categoria]['total'] += 1

        resposta = respostas_dict.get(nome_norm)
        if not resposta:
            militares_nao_informados.append(f"{dados['posto_grad']} {dados['nome_completo']}")
            continue

        status = str(resposta['status']).strip()
        posto_nome = f"{dados['posto_grad']} {dados['nome_completo']}"

        periodo = resposta.get('campo_outro')
        if periodo:
            posto_nome_saida = f"{posto_nome} {periodo}"
        else:
            posto_nome_saida = posto_nome

        if status.lower().find('presente') != -1 or status == "Presente":
            categorias_dados[categoria]['presentes'].append(posto_nome)
        else:
            categorias_dados[categoria]['afastamentos'].setdefault(status, []).append(posto_nome_saida)

    # ----------------------------
    # Gerar anúncio (COM ESPAÇO ENTRE TÓPICOS 🔹)
    # ----------------------------
    anuncio = f"""Bom dia!
Segue anúncio do dia

Anúncio CSC-PM
{data_formatada}

"""

    total_militares = 0
    total_civis = 0

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

    for categoria in ['OFICIAIS', 'PRAÇAS', 'CIVIS']:
        dados = categorias_dados[categoria]

        if categoria == 'CIVIS':
            total_civis = len(dados['presentes'])
        else:
            total_militares += len(dados['presentes'])

        anuncio += f"{categoria}\n"
        anuncio += "Efetivo total: \n"
        anuncio += f"🔸{dados['total']} - CSC-PM\n"

        # Presentes
        if dados['presentes']:
            anuncio += f"🔹{len(dados['presentes'])} Presentes:\n"
            for idx, nome in enumerate(dados['presentes'], 1):
                anuncio += f"    {idx}. {nome}\n"
            anuncio += "\n"  # <-- ESPAÇO ENTRE TÓPICOS 🔹

        # Afastamentos
        afast = dados['afastamentos']
        for status in sorted(afast.keys(), key=ordem_status):
            lista = afast[status]
            anuncio += f"🔹{len(lista)} {status}\n"
            for idx, info in enumerate(lista, 1):
                anuncio += f"    {idx}. {info}\n"
            anuncio += "\n"  # <-- ESPAÇO ENTRE TÓPICOS 🔹

        anuncio += "\n"

    anuncio += f"""Anúncio passado:
[PREENCHER MANUALMENTE]
[PREENCHER HORA]
➖➖➖➖➖ ➖ ➖
Efetivo presente:
Militares: {total_militares}
Civis: {total_civis}"""

    st.markdown("---")
    st.subheader("📢 ANÚNCIO GERADO:")
    st.code(anuncio, language='text')

    if militares_nao_informados:
        st.error("❌ MILITARES QUE NÃO RESPONDERAM:")
        for militar in militares_nao_informados:
            st.write(f"   • {militar}")
        st.warning(f"❌ Total de {len(militares_nao_informados)} militar(es) faltando no anúncio.")

    st.download_button(
        label="Baixar Anúncio de Presença",
        data=anuncio.encode('utf-8'),
        file_name="anuncio_presenca.txt",
        mime="text/plain"
    )

    st.success("✅ PROCESSO CONCLUÍDO!")
