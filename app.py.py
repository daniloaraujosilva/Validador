def extrair_dados_xml(root):
    ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
    infNFe = root.find('.//ns:infNFe', ns)
    if infNFe is None:
        return pd.DataFrame()
    ide = infNFe.find('ns:ide', ns)
    emit = infNFe.find('ns:emit', ns)
    dest = infNFe.find('ns:dest', ns)
    total = infNFe.find('ns:total/ns:ICMSTot', ns)
    det = infNFe.find('ns:det', ns)
    prod = det.find('ns:prod', ns)

    dados = {
        'tpNF': ide.findtext('ns:tpNF', default='', namespaces=ns),
        'Número_XML': ide.findtext('ns:nNF', default='', namespaces=ns),
        'Série_XML': ide.findtext('ns:serie', default='', namespaces=ns),
        'Data_Emissao_XML': ide.findtext('ns:dhEmi', default='', namespaces=ns),
        'Emitente_XML': emit.findtext('ns:xNome', default='', namespaces=ns),
        'CNPJ/CPF_XML': emit.findtext('ns:CNPJ', default='', namespaces=ns),
        'CFOP_XML': prod.findtext('ns:CFOP', default='', namespaces=ns),
        'Descrição_Produto_XML': prod.findtext('ns:xProd', default='', namespaces=ns),
        'Valor_XML': total.findtext('ns:vNF', default='', namespaces=ns),
        'Base_ICMS_XML_Total': total.findtext('ns:vBC', default='', namespaces=ns),
        'Valor_ICMS_XML_Total': total.findtext('ns:vICMS', default='', namespaces=ns),
        'Valor_vST': total.findtext('ns:vST', default='', namespaces=ns),
    }

    return pd.DataFrame([dados])


def corrigir_acentos(texto):
    try:
        return texto.encode("latin1").decode("utf-8")
    except:
        return texto
import streamlit as st
import streamlit.components.v1 as components


import pandas as pd
import xml.etree.ElementTree as ET
import re
import chardet
import unicodedata
import os

# ------------------------------------------------------------------
# Loader robusto de CFOP_Tabela produtos/ICMS
def _load_cfop_produtos():
    possiveis = ["cfop_produto_icms.xlsx", "CFOP_produtos.xlsx", "cfop_produtos.xlsx"]
    for fname in possiveis:
        if os.path.exists(fname):
            try:
                cfg = pd.read_excel(fname, dtype=str)
                cfg.columns = [c.strip().upper() for c in cfg.columns]
                if "CFOP_LIVRO" in cfg.columns:
                    return cfg["CFOP_LIVRO"].dropna().astype(str).tolist()
            except Exception:
                continue
    return []

# Lista global de CFOPs de entrada válidos
cfops_entrada_validos = _load_cfop_produtos()
# ------------------------------------------------------------------


def _to_float_safe(txt):
    txt = str(txt)
    if txt is None or txt.strip() == "":
        return 0.0
    # remove any dots used as thousand separator then replace comma with dot
    cleaned = re.sub(r"[^0-9,.-]", "", txt)
    if cleaned.count(",") == 1 and cleaned.count(".") > 0:
        cleaned = cleaned.replace(".", "")
    return float(cleaned.replace(",", ".") or 0)

def validar_soma_livro(row):
    try:
        base = _to_float_safe(row.get("BASE_CALCULO_Livro", "0"))
        isentas = _to_float_safe(row.get("ISENTAS_OU_N_TRIB_Livro", "0"))
        outras = _to_float_safe(row.get("OUTRAS_Livro", "0"))
        valor_contabil = _to_float_safe(row.get("VALOR_CONTÁBIL_Livro", "0"))
        soma = base + isentas + outras
        if abs(soma - valor_contabil) > 0.01:
            return "📉 Soma dos Valores do Livro Incorreta"
        return ""
    except:
        return ""

def gerar_status(row):
    motivos = []

    numero_livro = str(row.get("Número_Livro", "")).strip().lower()
    numero_xml = str(row.get("Número_XML", "")).strip().lower()

    if numero_livro in ["", "nan", "none"]:
        return "❌ XML não encontrado no Livro"
    if numero_xml in ["", "nan", "none"]:
        return "❌ Livro sem XML correspondente"

    if str(row.get("Divergência_Soma_Livro", "")).strip():
        motivos.append("📉 Soma Livro")
    if str(row.get("Divergência_CFOP_Equivalente", "")).strip():
        motivos.append("🔁 CFOP")
    if "💸 ICMS XML x Livro" in str(row.get("Divergência_ICMS_XML_Livro", "")):
        motivos.append("💸 ICMS XML x Livro")
    if str(row.get("Divergência_Valor_Data", "")).strip():
        motivos.extend(str(row["Divergência_Valor_Data"]).split(" | "))

    if motivos:
        return "⚠️ Divergência: " + " | ".join(motivos)
    return "✅ Validado"
def normalize(s: str) -> str:
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii").strip()
def formatar_cnpj_cpf(raw: str) -> str:
    nums = re.sub(r"\D", "", str(raw))
    if len(nums) == 14:
        return f"{nums[:2]}.{nums[2:5]}.{nums[5:8]}/{nums[8:12]}-{nums[12:]}"
    if len(nums) == 11:
        return f"{nums[:3]}.{nums[3:6]}.{nums[6:9]}-{nums[9:]}"
    return nums
def formatar_data_xml(data_str):
    try:
        return pd.to_datetime(data_str[:10], format="%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return ""
def formatar_data_livro(data_str):
    try:
        return pd.to_datetime(data_str, dayfirst=True).strftime("%d/%m/%Y")
    except Exception:
        return ""
def processar_livro(livro_file):
    df_raw = pd.read_excel(livro_file, dtype=str)
    df = df_raw.copy()

    # 🔍 Verificação robusta da coluna ISENTAS OU NÃO TRIBUTADAS
    colunas_livro = [c.strip().upper() for c in df_raw.columns]
    if not any("ISENTAS" in c and "TRIBUT" in c for c in colunas_livro):
        st.warning("⚠️ A coluna 'ISENTAS OU NÃO TRIBUTADAS' não foi encontrada no Livro Fiscal. Verifique o nome da coluna.")

    df.columns = [c.strip().upper() for c in df.columns]
    df = df.dropna(how="all")
    df = df[~df.apply(lambda r: r.astype(str).str.contains("FOLHA|LIVRO|REGISTRO|TOTAL", case=False).any(), axis=1)]
    df = df.rename(columns={
        "FISCAL": "CFOP_Livro",
        "ESPÉCIE":"ESPÉCIE","ESPECIE":"ESPÉCIE",
        "SÉRIE/\nSUB":"SÉRIE","SERIE":"SÉRIE",
        "NÚMERO":"NÚMERO","NUMERO":"NÚMERO",
        "VALOR\nCONTÁBIL":"VALOR_CONTÁBIL",
        "BASE DE CÁLCULO":"BASE_CALCULO",
        "IMPOSTO":"ICMS",
        "ISENTAS OU NÃO TRIBUTADAS":"ISENTAS_OU_N_TRIB",
        "\nOUTRAS":"OUTRAS"
    })
    if "EMITENTE" in df_raw.columns:
        df["Emitente_Livro"] = df_raw["EMITENTE"].astype(str)
        df["CNPJ/CPF_livro"] = (
            df_raw["EMITENTE"]
            .fillna("")
            .astype(str)
            .map(lambda x: x.split("-",1)[0])
            .map(formatar_cnpj_cpf)
        )
    else:
        df["CNPJ/CPF_livro"] = ""
        df["Emitente_Livro"] = ""
    
    # 🔁 Consolidação de CFOPs e somas para mesma nota (CNPJ + NÚMERO + SÉRIE)
    colunas_soma = ["BASE_CALCULO", "ICMS", "ISENTAS_OU_N_TRIB", "OUTRAS"]
    for col in colunas_soma:
        if col not in df.columns:
            df[col] = 0.0
        else:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
                .astype(float)
            )

    grupo_chave = ["CNPJ/CPF_livro", "NÚMERO", "SÉRIE"]
    soma_df = df.groupby(grupo_chave, as_index=False)[colunas_soma].sum()

    cfop_df = df.groupby(grupo_chave)["CFOP_Livro"] \
                .apply(lambda x: ", ".join(sorted(set(str(v) for v in x if pd.notna(v))))) \
                .reset_index()

    primeiro_df = df.drop_duplicates(subset=grupo_chave, keep="first").drop(columns=colunas_soma + ["CFOP_Livro"], errors="ignore")
    df = primeiro_df.merge(soma_df, on=grupo_chave).merge(cfop_df, on=grupo_chave).reset_index(drop=True)

    df["NÚMERO"] = df["NÚMERO"].fillna("").astype(str).map(normalize)
    df["SÉRIE"]  = df["SÉRIE"].fillna("").astype(str).map(normalize)
    df["ESPÉCIE_CLEAN"] = (
        df["ESPÉCIE"]
        .str.replace(r"[^A-Za-z0-9]", "", regex=True)
        .str.upper()
        .map({"NFE":"55","CTE":"57","NFC":"65"})
        .fillna(df["ESPÉCIE"])
    )
    default = df["SÉRIE"].mode().iloc[0] if not df["SÉRIE"].mode().empty else "1"
    df["SÉRIE"] = df["SÉRIE"].replace("", default)
    return df
@st.cache_data(show_spinner="⏳ Processando XMLs...")
def extrair_documentos(xml_files):
    registros = []
    ns_nfe = {"nfe": "http://www.portalfiscal.inf.br/nfe"}
    ns_cte = {"cte": "http://www.portalfiscal.inf.br/cte"}
    for f in xml_files:
        nome_arquivo = f.name
        f.seek(0)
        raw = f.read()
        enc = chardet.detect(raw)["encoding"] or "utf-8"
        xml = raw.decode(enc, errors="ignore")
        xml = re.sub(r"[^ -~\n\r\t]", "", xml)
        root = ET.fromstring(xml)
        # Tenta CT-e primeiro
        ide_cte = root.find(".//cte:ide", namespaces=ns_cte)
        if ide_cte is not None:
            mod = "57"
            emit_cte = root.find(".//cte:emit", namespaces=ns_cte)
            num = normalize(ide_cte.findtext("cte:nCT", "", namespaces=ns_cte))
            ser = normalize(ide_cte.findtext("cte:serie", "", namespaces=ns_cte))
            tp = ide_cte.findtext("cte:tpEmis", "1", namespaces=ns_cte)
            demissao = ide_cte.findtext("cte:dhEmi", namespaces=ns_cte)
            doc = emit_cte.findtext("cte:CNPJ", namespaces=ns_cte) or emit_cte.findtext("cte:CPF", namespaces=ns_cte)
            valor = root.findtext(".//cte:vTPrest", namespaces=ns_cte) or "0"
            cfops = [normalize(tag.text) for tag in root.findall(".//cte:CFOP", namespaces=ns_cte) if tag.text]
            registros.append({
            "Arquivo_XML": nome_arquivo,
            "Tipo": "CT-e",
                "Número": num,
                "Série": ser,
                "Espécie": mod,
                "tpNF": tp,
                "Emitente_XML": normalize(emit_cte.findtext("cte:xNome", namespaces=ns_cte) or ""),
                "CNPJ/CPF_xml": formatar_cnpj_cpf(doc),
                "CFOP_XML": ", ".join(sorted(set(cfops))),
                "Valor_XML": valor,
                "Valor_vST": "0",
                "Data_XML": demissao
            })
            continue  # Se for CT-e, não tenta processar como NF-e
        # Tenta NF-e se não for CT-e
        ide_nfe = root.find(".//nfe:ide", namespaces=ns_nfe)
        if ide_nfe is not None:
            mod = normalize(ide_nfe.findtext("nfe:mod", "", namespaces=ns_nfe))
            if mod != "55":
                continue  # Só aceita NF-e modelo 55
            num = normalize(ide_nfe.findtext("nfe:nNF", "", namespaces=ns_nfe))
            ser = normalize(ide_nfe.findtext("nfe:serie", "", namespaces=ns_nfe))
            tp = ide_nfe.findtext("nfe:tpNF", "1", namespaces=ns_nfe)
            demissao = ide_nfe.findtext("nfe:dhEmi", namespaces=ns_nfe) or ide_nfe.findtext("nfe:dEmi", namespaces=ns_nfe) or ""
            emit = root.find(".//nfe:emit", namespaces=ns_nfe)
            doc = emit.findtext("nfe:CNPJ", namespaces=ns_nfe) or emit.findtext("nfe:CPF", namespaces=ns_nfe) or ""
            valor = root.findtext(".//nfe:vNF", namespaces=ns_nfe)or "0"
            vst = root.findtext(".//nfe:ICMSTot/nfe:vST", namespaces=ns_nfe) or "0"
            cfop_tags = root.findall(".//nfe:CFOP", namespaces=ns_nfe)
            cfops = sorted(set(normalize(tag.text) for tag in cfop_tags if tag.text))
            registros.append({
            "Arquivo_XML": nome_arquivo,
            "Tipo": "NF-e",
                "Número": num,
                "Série": ser,
                "Espécie": mod,
                "tpNF": tp,
                "Emitente_XML": normalize(emit.findtext("nfe:xNome", namespaces=ns_nfe) or ""),
                "CNPJ/CPF_xml": formatar_cnpj_cpf(doc),
                "CFOP_XML": ", ".join(cfops),
                "Valor_XML": valor,
                "Valor_vST": vst,
                "Data_XML": demissao,
            })
    return pd.DataFrame(registros)


def extrair_icms_cte(xml_files):
    registros = []
    ns_cte = {"cte": "http://www.portalfiscal.inf.br/cte"}
    for f in xml_files:
        nome_arquivo = f.name
        try:
            f.seek(0)
            raw = f.read()
            enc = chardet.detect(raw)["encoding"] or "utf-8"
            xml = raw.decode(enc, errors="ignore")
            xml = re.sub(r"[^ -~\n\r\t]", "", xml)
            root = ET.fromstring(xml)
            ide = root.find(".//cte:ide", namespaces=ns_cte)
            emit = root.find(".//cte:emit", namespaces=ns_cte)
            icms_tag = root.find(".//cte:ICMS", namespaces=ns_cte)
            if ide is None or icms_tag is None:
                continue
            mod = "57"
            cst = icms_tag.findtext(".//cte:CST", "", namespaces=ns_cte)
            if cst not in ["00", "20"]:
                continue
            vbc = icms_tag.findtext(".//cte:vBC", "0", namespaces=ns_cte)
            vicms = icms_tag.findtext(".//cte:vICMS", "0", namespaces=ns_cte)
            num = normalize(ide.findtext("cte:nCT", "", namespaces=ns_cte))
            ser = normalize(ide.findtext("cte:serie", "", namespaces=ns_cte))
            cnpj = formatar_cnpj_cpf(emit.findtext("cte:CNPJ", namespaces=ns_cte) or "")
            registros.append({
                "Número": num,
                "Série": ser,
                "Espécie": mod,
                "CNPJ/CPF_xml": cnpj,
                "Base_ICMS_XML_Total": float(vbc.replace(",", ".")),
                "Valor_ICMS_XML_Total": float(vicms.replace(",", ".")),
                "Descrição_Produto_XML": "Serviço CT-e",
                "CST": cst
            })
        except:
            continue
    return pd.DataFrame(registros)

# O restante do código continua no próximo bloco, por tamanho.
@st.cache_data(show_spinner="⏳ Processando XMLs...")
def extrair_icms_produtos(xml_files):
    registros = []
    ns = {
        "nfe": "http://www.portalfiscal.inf.br/nfe",
        "det": "http://www.portalfiscal.inf.br/nfe"
    }
    for f in xml_files:
        nome_arquivo = f.name
        try:
            f.seek(0)
            raw = f.read()
            enc = chardet.detect(raw)["encoding"] or "utf-8"
            xml = raw.decode(enc, errors="ignore")
            xml = re.sub(r"[^ -~\n\r\t]", "", xml)
            root = ET.fromstring(xml)
            chave = root.findtext(".//nfe:infProt/nfe:chNFe", namespaces=ns) or ""
            ide = root.find(".//nfe:ide", namespaces=ns)
            emit = root.find(".//nfe:emit", namespaces=ns)
            mod = normalize(ide.findtext("nfe:mod", "", namespaces=ns))
            if mod != "55":
                continue  # Só para NF-e
            num = normalize(ide.findtext("nfe:nNF", "", namespaces=ns))
            ser = normalize(ide.findtext("nfe:serie", "", namespaces=ns))
            esp = mod
            cnpj = formatar_cnpj_cpf(emit.findtext("nfe:CNPJ", namespaces=ns) or "")
            produtos = []
            total_vbc = 0.0
            total_vicms = 0.0
            cst_validado = False
            for det in root.findall(".//nfe:det", namespaces=ns):
                prod = det.find("nfe:prod", namespaces=ns)
                imposto = det.find("nfe:imposto/nfe:ICMS", namespaces=ns)
                if prod is not None:
                    xprod = prod.findtext("nfe:xProd", "", namespaces=ns)
                    produtos.append(xprod)
                if imposto is not None:
                    icms_tag = next(iter(imposto), None)
                    if icms_tag is not None:
                        cst = icms_tag.findtext("nfe:CST", "", namespaces=ns)
                        if cst in ["00", "20"]:
                            cst_validado = True
                            vbc = icms_tag.findtext("nfe:vBC", "0", namespaces=ns)
                            vicms = icms_tag.findtext("nfe:vICMS", "0", namespaces=ns)
                            try:
                                total_vbc += float(vbc.replace(",", "."))
                                total_vicms += float(vicms.replace(",", "."))
                            except:
                                pass
            if cst_validado:
                registros.append({
                    "Número": num,
                    "Série": ser,
                    "Espécie": esp,
                    "CNPJ/CPF_xml": cnpj,
                    "Descrição_Produto_XML": " | ".join(produtos),
                    "Base_ICMS_XML_Total": total_vbc,
                    "Valor_ICMS_XML_Total": total_vicms
                })
        except Exception as e:
            continue
    return pd.DataFrame(registros)


def extrair_icms_cte(xml_files):
    registros = []
    ns_cte = {"cte": "http://www.portalfiscal.inf.br/cte"}
    for f in xml_files:
        nome_arquivo = f.name
        try:
            f.seek(0)
            raw = f.read()
            enc = chardet.detect(raw)["encoding"] or "utf-8"
            xml = raw.decode(enc, errors="ignore")
            xml = re.sub(r"[^ -~\n\r\t]", "", xml)
            root = ET.fromstring(xml)
            ide = root.find(".//cte:ide", namespaces=ns_cte)
            emit = root.find(".//cte:emit", namespaces=ns_cte)
            icms_tag = root.find(".//cte:ICMS", namespaces=ns_cte)
            if ide is None or icms_tag is None:
                continue
            mod = "57"
            cst = icms_tag.findtext(".//cte:CST", "", namespaces=ns_cte)
            if cst not in ["00", "20"]:
                continue
            vbc = icms_tag.findtext(".//cte:vBC", "0", namespaces=ns_cte)
            vicms = icms_tag.findtext(".//cte:vICMS", "0", namespaces=ns_cte)
            num = normalize(ide.findtext("cte:nCT", "", namespaces=ns_cte))
            ser = normalize(ide.findtext("cte:serie", "", namespaces=ns_cte))
            cnpj = formatar_cnpj_cpf(emit.findtext("cte:CNPJ", namespaces=ns_cte) or "")
            registros.append({
                "Número": num,
                "Série": ser,
                "Espécie": mod,
                "CNPJ/CPF_xml": cnpj,
                "Base_ICMS_XML_Total": float(vbc.replace(",", ".")),
                "Valor_ICMS_XML_Total": float(vicms.replace(",", ".")),
                "Descrição_Produto_XML": "Serviço CT-e",
                "CST": cst
            })
        except:
            continue
    return pd.DataFrame(registros)

st.set_page_config(page_title="Validador Fiscal", layout="wide")

# >>> INÍCIO: Customização de Cor do Botão de Download >>>
st.markdown(
    """
    <style>
      div.stDownloadButton > button {
        background-color: #0e4c92 !important;
        color: #ffffff !important;
      }
      div.stDownloadButton > button:hover {
        background-color: #073d77 !important;
      }
      div.stDownloadButton svg {
        fill: #ffffff !important;
      }
    </style>
    """,
    unsafe_allow_html=True,
)
# <<< FIM: Customização de Cor do Botão de Download <<<



import streamlit as st
import streamlit.components.v1 as components
import base64
import os

# Inserção do logo da empresa (centralizado e responsivo)
if os.path.exists("logo_empresa.png"):
    st.markdown(
        f"""
        <div style='text-align: center; padding: 10px;'>
            <img src="data:image/png;base64,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" 
                 style="width:200px; border-radius: 10px;" />
        </div>
        """,
        unsafe_allow_html=True
    )


st.markdown("""
    <style>
        .block-container { padding-top: 2rem; }
        h1 { color: #0e4c92; }
        .dataframe { border: 1px solid #ccc; border-radius: 10px; }
        .stSubheader { margin-top: 2rem; color: #1a1a1a; }
        .stDownloadButton {
            background-color: #0e4c92;
            color: white;
            border-radius: 5px;
        }
        .stDownloadButton:hover {
            background-color: #073d77;
        }
    </style>
""", unsafe_allow_html=True)


st.markdown("""
    <div style='text-align: center; padding: 10px 0; border-bottom: 2px solid #ccc;'>
        <h1 style='color:#0e4c92;'>📊 Validador Fiscal XML x Livro</h1>
        <p style='color:#555;'>Análise cruzada NF-e / CT-e com Livro Fiscal</p>
    </div>
""", unsafe_allow_html=True)



left, col_file, col_xml, right = st.columns([1, 3, 3, 1])
with col_file:
    livro_file = st.file_uploader("📘 Livro Fiscal", type=["xls", "xlsx"])
with col_xml:
    xml_files = uploaded_files = st.file_uploader("🧾 XMLs (NF‑e / CT‑e)", type=["xml"], accept_multiple_files=True)
if livro_file and xml_files:
    df_livro_original = pd.read_excel(livro_file, dtype=str)
    df_livro = processar_livro(livro_file)
    df_livro = df_livro.applymap(lambda x: corrigir_acentos(x) if isinstance(x, str) else x)
    df_xml = extrair_documentos(xml_files)
    df_xml = df_xml.applymap(lambda x: corrigir_acentos(x) if isinstance(x, str) else x)

    st.success("✅ Arquivos carregados com sucesso. Processando...")

    df_norm = df_xml[df_xml["tpNF"] != "0"].copy()
    df_dev  = df_xml[df_xml["tpNF"] == "0"].copy()

    df_jnorm = pd.merge(
        df_norm,
        df_livro, how="left",
        left_on=["Número", "Série", "Espécie", "CNPJ/CPF_xml"],
        right_on=["NÚMERO", "SÉRIE", "ESPÉCIE_CLEAN", "CNPJ/CPF_livro"],
        suffixes=("", "_livro")
    )

    df_jdev = pd.merge(
        df_dev,
        df_livro, how="left",
        left_on=["Número", "Série", "Espécie"],
        right_on=["NÚMERO", "SÉRIE", "ESPÉCIE_CLEAN"],
        suffixes=("", "_livro")
    )

    df_final = pd.concat([df_jnorm, df_jdev], ignore_index=True)

    if "CFOP_XML" not in df_final.columns:
        df_final["CFOP_XML"] = ""
    if "CFOP_Livro" not in df_final.columns:
        df_final["CFOP_Livro"] = ""

    colunas_renomeadas = {
        "Tipo": "Tipo_XML",
        "Número": "Número_XML",
        "Série": "Série_XML",
        "Espécie": "Espécie_XML",
        "tpNF": "tpNF_XML",
        "CNPJ/CPF_xml": "CNPJ/CPF_XML",
        "NÚMERO": "Número_Livro",
        "SÉRIE": "Série_Livro",
        "ESPÉCIE": "Espécie_Livro",
        "ESPÉCIE_CLEAN": "Espécie_Livro_CLEAN",
        "CNPJ/CPF_livro": "CNPJ/CPF_Livro",
        "Emitente_XML": "Emitente_XML",
        "Emitente_Livro": "Emitente_Livro"
    }

    df_final_renomeado = df_final.rename(columns=colunas_renomeadas)
    df_final_renomeado["Status"] = df_final_renomeado.apply(gerar_status, axis=1)
    # ─── ALERTA ICMS-ST (na visão final) ───
    if "Valor_vST" in df_final_renomeado.columns:
        df_final_renomeado["Valor_vST"] = df_final_renomeado["Valor_vST"].astype(float)
        df_final_renomeado["Alerta_ICMS_ST"] = df_final_renomeado["Valor_vST"].apply(lambda x: "📌 **Anexar Guia e Comprovante de ST.**" if x > 0 else "")



    if df_livro.empty:
        st.warning("⚠️ Livro Fiscal está vazio ou não possui colunas compatíveis.")
    if df_xml.empty:
        st.warning("⚠️ Nenhum dado foi extraído dos XMLs enviados.")
        st.warning("⚠️ Nenhum dado foi extraído dos XMLs enviados.")
    df_norm = df_xml[df_xml["tpNF"] != "0"].copy()
    df_dev  = df_xml[df_xml["tpNF"] == "0"].copy()
    df_jnorm = pd.merge(
        df_norm,
        df_livro, how="left",
        left_on=["Número","Série","Espécie","CNPJ/CPF_xml"],
        right_on=["NÚMERO","SÉRIE","ESPÉCIE_CLEAN","CNPJ/CPF_livro"],
        suffixes=("","_livro")
    )
    df_jdev = pd.merge(
        df_dev, df_livro, how="left",
        left_on=["Número","Série","Espécie"],
        right_on=["NÚMERO","SÉRIE","ESPÉCIE_CLEAN"],
        suffixes=("","_livro")
    )
    df_final = pd.concat([df_jnorm, df_jdev], ignore_index=True)
    df_final["CNPJ/CPF"] = df_final["CNPJ/CPF_xml"].mask(df_final["CNPJ/CPF_xml"] == "", df_final["CNPJ/CPF_livro"])
    df_final["Emitente_Livro"] = df_final["Emitente_Livro"].astype(str).apply(lambda x: x.split(" - ")[-1].strip() if " - " in x else x)
    if "CFOP_XML" not in df_final.columns:
        df_final["CFOP_XML"] = ""
    if "CFOP_Livro" not in df_final.columns:
        df_final["CFOP_Livro"] = ""
    colunas_renomeadas = {
        "Tipo": "Tipo_XML",
        "Número": "Número_XML",
        "Série": "Série_XML",
        "Espécie": "Espécie_XML",
        "tpNF": "tpNF_XML",
        "CNPJ/CPF_xml": "CNPJ/CPF_XML",
        "NÚMERO": "Número_Livro",
        "SÉRIE": "Série_Livro",
        "ESPÉCIE": "Espécie_Livro",
        "ESPÉCIE_CLEAN": "Espécie_Livro_CLEAN",
        "CNPJ/CPF_livro": "CNPJ/CPF_Livro",
        "VALOR_CONTÁBIL": "Valor_XML",
        "VALOR_CONTÁBIL": "VALOR_CONTÁBIL_Livro",
        "Valor_XML": "Valor_XML",
        "BASE_CALCULO": "BASE_CALCULO_Livro",
        "ICMS": "ICMS_Livro",
        "ISENTAS_OU_N_TRIB": "ISENTAS_OU_N_TRIB_Livro",
        "OUTRAS": "OUTRAS_Livro",
        "DOC.": "DATA_EMISSAO_Livro",
        "Emitente_XML": "Emitente_XML",
        "CFOP_XML": "CFOP_XML",
        "Valor_XML": "Valor_XML",
        "Emitente_Livro": "Emitente_Livro"
    }
    df_final_renomeado = df_final.rename(columns=colunas_renomeadas)
    df_final_renomeado["Status"] = df_final_renomeado.apply(gerar_status, axis=1)
    # ─── ALERTA ICMS-ST (na visão final) ───
    if "Valor_vST" in df_final_renomeado.columns:
        df_final_renomeado["Valor_vST"] = df_final_renomeado["Valor_vST"].astype(float)
        df_final_renomeado["Alerta_ICMS_ST"] = df_final_renomeado["Valor_vST"].apply(lambda x: "📌 **Anexar Guia e Comprovante de ST.**" if x > 0 else "")

    # ------------------------------------------------------------------
    # Remover duplicidades geradas por múltiplas linhas no Livro
    chave_uni = ["Número_XML", "Série_XML", "Espécie_XML", "CNPJ/CPF_XML"]
    df_final_renomeado = (
        df_final_renomeado
            .drop_duplicates(subset=chave_uni, keep="first")
            .reset_index(drop=True)
    )
    # ------------------------------------------------------------------

    # ------------------------------------------------------------------
    # Garantir 1 linha por nota (impede duplicidades quando o Livro possui
    # várias linhas para o mesmo documento)
    chave_uni = ["Número_XML", "Série_XML", "Espécie_XML", "CNPJ/CPF_XML"]
    df_final_renomeado = (
        df_final_renomeado
            .drop_duplicates(subset=chave_uni, keep="first")
            .reset_index(drop=True)
    )
    # ------------------------------------------------------------------
    # >>> INÍCIO: Adicionar notas do Livro sem XML correspondente >>>
    chaves_livro = ['NÚMERO','SÉRIE','ESPÉCIE_CLEAN','CNPJ/CPF_livro']
    df_livro_sem_xml = df_livro[
        ~df_livro[chaves_livro].apply(tuple, axis=1).isin(
            df_final_renomeado[['Número_Livro','Série_Livro','Espécie_Livro_CLEAN','CNPJ/CPF_Livro']]
            .apply(tuple, axis=1)
        )
    ].copy()
    if not df_livro_sem_xml.empty:
        df_livro_sem_xml.rename(columns={
            'NÚMERO':'Número_Livro',
            'SÉRIE':'Série_Livro',
            'ESPÉCIE_CLEAN':'Espécie_Livro_CLEAN',
            'CNPJ/CPF_livro':'CNPJ/CPF_Livro'
        }, inplace=True)
        for col in ['Número_XML','Série_XML','Espécie_XML','tpNF_XML','CNPJ/CPF_XML','Emitente_XML','CFOP_XML','Valor_XML','Data_XML']:
            df_livro_sem_xml[col] = ''
        df_livro_sem_xml['Status'] = '❌ Livro sem XML correspondente'
        df_final_renomeado = pd.concat([df_final_renomeado, df_livro_sem_xml], ignore_index=True)
    # <<< FIM: Adicionar notas do Livro sem XML correspondente <<<
    df_final_renomeado["Data_Emissao_XML"] = df_final_renomeado["Data_XML"].apply(formatar_data_xml)
    
    if "DATA_EMISSAO_Livro" not in df_final_renomeado.columns and "DOC." in df_final.columns:
        df_final_renomeado["DATA_EMISSAO_Livro"] = df_final["DOC."]
    df_final_renomeado["Data_Emissao_Livro"] = df_final_renomeado["DATA_EMISSAO_Livro"].apply(formatar_data_livro)
    try:
        tabela_cfop = pd.read_excel("cfop_tabela.xlsx", dtype=str)
        tabela_cfop.columns = [c.strip().upper() for c in tabela_cfop.columns]
        # Carregar mapa de descrições de CFOP como strings
        tabela_cfop = pd.read_excel("cfop_tabela.xlsx", dtype=str)
        tabela_cfop.columns = [c.strip().upper() for c in tabela_cfop.columns]
        mapa_descricoes = {
            str(cfop).strip(): str(desc).strip()
            for cfop, desc in zip(
                tabela_cfop["CFOP"],
                tabela_cfop["DESCRIÇÃO RESUMIDA"]
            )
        }
        def safe_concat_cfop(x):
            codes = [c.strip() for c in str(x).split(",") if c.strip()]
            descs = [mapa_descricoes.get(code, "CFOP não encontrado") for code in codes]
            return ", ".join(str(d) for d in descs)
        # Aplicar descrições seguras para evitar TypeError

        # Carregar equivalências de CFOP no mesmo escopo
        equivalencias_df = pd.read_excel("cfop_equivalencias.xlsx", dtype=str)
        equivalencias_dict = {}
        for _, r in equivalencias_df.iterrows():
            xml_cfop = str(r["CFOP_XML"]).strip()
            livr_cfop = str(r["CFOP_LIVRO_EQUIVALENTE"]).strip()
            equivalencias_dict.setdefault(xml_cfop, []).append(livr_cfop)

        def validar_equivalencia_cfop(row):
            xml_codes = [c.strip() for c in str(row["CFOP_XML"]).split(",") if c.strip()]
            livro_codes = [c.strip() for c in str(row["CFOP_Livro"]).split(",") if c.strip()]
            for x in xml_codes:
                if not set(equivalencias_dict.get(x, [])).intersection(livro_codes):
                    return "❌ CFOP não equivalente ao XML"
            return ""
        
        df_final_renomeado["CFOP_XML_Descrição"] = df_final_renomeado["CFOP_XML"].apply(safe_concat_cfop)
        df_final_renomeado["CFOP_Livro_Descrição"] = df_final_renomeado["CFOP_Livro"].apply(safe_concat_cfop)
        df_final_renomeado["Divergência_CFOP_Equivalente"] = df_final_renomeado.apply(validar_equivalencia_cfop, axis=1)
    except Exception as e:
        st.warning(f"Erro ao validar CFOP por equivalência: {e}")
    
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO


def limpar_valor_icms(valor):
    if pd.isna(valor) or str(valor).strip() in ["", "-", "nan"]:
        return Decimal("0.00")
    try:
        val = re.sub(r"[^\d,\.]", "", str(valor))
        if val.count(",") == 1 and val.count(".") > 0:
            val = val.replace(".", "")
        val = val.replace(",", ".")
        return Decimal(val)
    except Exception:
        return Decimal("0.00")

def validar_icms_xml_livro(row):
    try:
        val_livro = limpar_valor_icms(row.get("ICMS_Livro", "0"))
        val_xml = limpar_valor_icms(row.get("Valor_ICMS_XML_Total", "0"))

        if row.get("Espécie_XML") == "57" and val_livro == 0:
            return ""  # Isenta para CT-e se zero

        dif = abs(val_xml - val_livro).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if dif > Decimal("0.01"):
            return "💸 ICMS XML x Livro"
        return ""
    except (InvalidOperation, ValueError, TypeError):
        return "❌ Erro ICMS na validação"
def validar_cnpj_xml_livro(row):
    cnpj_xml = str(row.get("CNPJ/CPF_XML", "")).strip()
    cnpj_livro = str(row.get("CNPJ/CPF_Livro", "")).strip()
    tipo_nota = str(row.get("tpNF_XML", "1")).strip()
    if tipo_nota == "0" and cnpj_xml == cnpj_livro:
        return ""  # Devolução própria
    if cnpj_xml and cnpj_livro and cnpj_xml != cnpj_livro:
        return "❌ CNPJ divergente entre XML e Livro"
    return ""

    cnpj_xml = row.get("CNPJ/CPF_XML", "").strip()
    cnpj_livro = row.get("CNPJ/CPF_Livro", "").strip()
    tipo_nota = str(row.get("tpNF_XML", "1")).strip()
    if tipo_nota == "0" and cnpj_xml == cnpj_livro:
        return ""  # Devolução própria
    if cnpj_xml and cnpj_livro and cnpj_xml != cnpj_livro:
        return "❌ CNPJ divergente entre XML e Livro"
    return ""



def validar_data_valor(row):
    try:
        # Conversão segura do Valor Contábil (Livro) em formato BR
        vl = str(row.get("VALOR_CONTÁBIL_Livro", "0"))
        vl = vl.replace(".", "").replace(",", ".")
        valor_livro = float(vl)

        # Conversão dinâmica do Valor (XML), BR ou INT
        vx = str(row.get("Valor_XML", "0"))
        if "," in vx:
            vx = vx.replace(".", "").replace(",", ".")
        valor_xml = float(vx)

        # Verificação de divergência de valores
        if abs(valor_xml - valor_livro) > 0.01:
            return "💰 Valor"
        return ""
    except Exception:
        return "💰 Valor"

def gerar_correcao(status:str) -> str:
    """Gera texto de correções agregando todas as divergências presentes no status."""
    correcoes=[]
    if "XML não encontrado no Livro" in status:
        correcoes.append("Revisar se a nota foi escriturada corretamente no Livro Fiscal.")
    if "Livro sem XML correspondente" in status:
        correcoes.append("Verificar se o XML da nota fiscal foi disponibilizado corretamente.")
    if "Soma Livro" in status:
        correcoes.append("Revisar valores de base de cálculo, isentas e outras no Livro.")
    if "CFOP" in status:
        correcoes.append("Verificar se o CFOP utilizado está de acordo com o XML ou equivalência cadastrada.")
    if "ICMS XML x Livro" in status:
        correcoes.append("Revisar valores de ICMS declarados no Livro e no XML.")
    if "Data" in status:
        correcoes.append("Corrigir data de emissão conforme o XML.")
    if "Valor" in status:
        correcoes.append("Revisar valor contábil informado no Livro comparado ao XML.")
    if not correcoes and "OK" in status:
        correcoes.append("✅ Validado")
    return " | ".join(correcoes)

try:
#     cfop_cfg = pd.read_excel("cfop_produto_icms.xlsx", dtype=str)
#     cfops_entrada_validos = cfop_cfg["CFOP_LIVRO"].dropna().astype(str).tolist()

    df_icms_nfe = extrair_icms_produtos(xml_files)
    df_icms_cte = extrair_icms_cte(xml_files)
    df_icms = pd.concat([df_icms_nfe, df_icms_cte], ignore_index=True)

    def sanitize_valor(valor):
        try:
            return float(re.sub(r"[^\d,\.]", "", str(valor)).replace(",", "."))
        except:
            return 0.0

    def aplicar_regras_extracao(row):
        icms_livro = sanitize_valor(row.get("ICMS_Livro", "0"))
        especie = row.get("Espécie_XML", "")
        if especie == "57":
            if icms_livro <= 0:
                return pd.Series(["", "", ""])
        if especie == "57" or row.get("CFOP_Livro", "") in cfops_entrada_validos:
            filtro = (
                (df_icms["Número"] == row["Número_XML"]) &
                (df_icms["Série"] == row["Série_XML"]) &
                (df_icms["Espécie"] == row["Espécie_XML"]) &
                (df_icms["CNPJ/CPF_xml"] == row["CNPJ/CPF_XML"])
            )
            resultado = df_icms[filtro]
            if not resultado.empty:
                dados = resultado.iloc[0]
                return pd.Series([
                    dados.get("Descrição_Produto_XML", ""),
                    dados.get("Base_ICMS_XML_Total", ""),
                    dados.get("Valor_ICMS_XML_Total", "")
                ])
        return pd.Series(["", "", ""])

    df_final_renomeado[["Descrição_Produto_XML", "Base_ICMS_XML_Total", "Valor_ICMS_XML_Total"]] = df_final_renomeado.apply(aplicar_regras_extracao, axis=1)
    df_final_renomeado["Divergência_ICMS_XML_Livro"] = df_final_renomeado.apply(validar_icms_xml_livro, axis=1)
    df_final_renomeado["Divergência_Valor_Data"] = df_final_renomeado.apply(validar_data_valor, axis=1)
    df_final_renomeado["Divergência_Soma_Livro"] = df_final_renomeado.apply(validar_soma_livro, axis=1)
    df_final_renomeado["Status"] = df_final_renomeado.apply(gerar_status, axis=1)
    # ─── ALERTA ICMS-ST (na visão final) ───
    if "Valor_vST" in df_final_renomeado.columns:
        df_final_renomeado["Valor_vST"] = df_final_renomeado["Valor_vST"].astype(float)
        df_final_renomeado["Alerta_ICMS_ST"] = df_final_renomeado["Valor_vST"].apply(lambda x: "📌 **Anexar Guia e Comprovante de ST.**" if x > 0 else "")

    # Criar coluna Correção com todas as recomendações pertinentes
    df_final_renomeado["Correção"] = df_final_renomeado["Status"].apply(gerar_correcao)

    # ✅ Reformatar campo Valor_XML para padrão brasileiro na visualização
    if "Valor_XML" in df_final_renomeado.columns:
        df_final_renomeado["Valor_XML"] = df_final_renomeado["Valor_XML"].apply(
            lambda x: "{:,.2f}".format(float(str(x).replace(",", "."))).replace(",", "v").replace(".", ",").replace("v", ".")
            if str(x).replace(",", ".").replace(".", "").isdigit() else x
        )


    if "BASE_CALCULO_Livro" in df_final_renomeado.columns:
        df_final_renomeado["BASE_CALCULO_Livro"] = df_final_renomeado["BASE_CALCULO_Livro"].apply(
            lambda x: "{:,.2f}".format(float(str(x).replace(",", "."))).replace(",", "v").replace(".", ",").replace("v", ".")
            if str(x).replace(",", ".").replace(".", "").isdigit() else x
        )

    if "ISENTAS_OU_N_TRIB_Livro" in df_final_renomeado.columns:
        df_final_renomeado["ISENTAS_OU_N_TRIB_Livro"] = df_final_renomeado["ISENTAS_OU_N_TRIB_Livro"].apply(
            lambda x: "{:,.2f}".format(float(str(x).replace(",", "."))).replace(",", "v").replace(".", ",").replace("v", ".")
            if str(x).replace(",", ".").replace(".", "").isdigit() else x
        )

    if "OUTRAS_Livro" in df_final_renomeado.columns:
        df_final_renomeado["OUTRAS_Livro"] = df_final_renomeado["OUTRAS_Livro"].apply(
            lambda x: "{:,.2f}".format(float(str(x).replace(",", "."))).replace(",", "v").replace(".", ",").replace("v", ".")
            if str(x).replace(",", ".").replace(".", "").isdigit() else x
        )

    if "ICMS_Livro" in df_final_renomeado.columns:
        df_final_renomeado["ICMS_Livro"] = df_final_renomeado["ICMS_Livro"].apply(
            lambda x: "{:,.2f}".format(float(str(x).replace(",", "."))).replace(",", "v").replace(".", ",").replace("v", ".")
            if str(x).replace(",", ".").replace(".", "").isdigit() else x
        )

    if "Base_ICMS_XML_Total" in df_final_renomeado.columns:
        df_final_renomeado["Base_ICMS_XML_Total"] = df_final_renomeado["Base_ICMS_XML_Total"].apply(
            lambda x: "{:,.2f}".format(float(str(x).replace(",", "."))).replace(",", "v").replace(".", ",").replace("v", ".")
            if str(x).replace(",", ".").replace(".", "").isdigit() else x
        )

    if "Valor_ICMS_XML_Total" in df_final_renomeado.columns:
        df_final_renomeado["Valor_ICMS_XML_Total"] = df_final_renomeado["Valor_ICMS_XML_Total"].apply(
            lambda x: "{:,.2f}".format(float(str(x).replace(",", "."))).replace(",", "v").replace(".", ",").replace("v", ".")
            if str(x).replace(",", ".").replace(".", "").isdigit() else x
        )


    # 🔢 Normalizar visualmente colunas com números e séries (formato sem zeros à esquerda)
    colunas_formatar = ["Número_Livro", "Número_XML", "Série_Livro", "Série_XML", "Espécie_Livro", "CFOP_Livro", "CFOP_XML"]
    for col in colunas_formatar:
        if col in df_final_renomeado.columns:
            df_final_renomeado[col] = df_final_renomeado[col].astype(str).str.lstrip("0")


    # 🔢 Normalizar visualmente colunas com números e séries (formato sem zeros à esquerda)
    colunas_formatar = ["Número_Livro", "Número_XML", "Série_Livro", "Série_XML", "Espécie_Livro", "CFOP_Livro", "CFOP_XML"]
    for col in colunas_formatar:
        if col in df_final_renomeado.columns:
            df_final_renomeado[col] = df_final_renomeado[col].astype(str).str.lstrip("0")



    df_final_renomeado["Divergência"] = df_final_renomeado["Status"]
    df_final_renomeado["Divergência_CNPJ"] = df_final_renomeado.apply(validar_cnpj_xml_livro, axis=1)

    # Mapeamento de nomes para exibição clara (XML vs Livro)

    # Métricas de contagem após todas as divergências
    total_xml = len(df_xml)
    total_livro = len(df_livro)
    total_divergentes = df_final_renomeado[df_final_renomeado['Status'] != '✅ Validado'].shape[0]
    taxa_conformidade = 100 * (1 - (total_divergentes / max(len(df_final_renomeado), 1)))
    colk1, colk2, colk3, colk4 = st.columns(4)
    colk1.metric("🧾 XMLs Carregados", f"{total_xml}")
    colk2.metric("📘 Registros no Livro", f"{total_livro}")
    colk3.metric("⚠️ Divergências", f"{total_divergentes}")
    colk4.metric("📈 % Conformidade", f"{taxa_conformidade:.1f}%")

    display_names = {
    'Valor_vST': 'Valor ST (XML)',
    'Alerta_ICMS_ST': 'Alerta ST',
        "Número_Livro": "Número (Livro)",
        "Número_XML": "Número (XML)",
        "Série_Livro": "Série (Livro)",
        "Série_XML": "Série (XML)",
        "Espécie_Livro": "Modelo (Livro)",
        "Espécie_XML": "Modelo (XML)",
        "Emitente_Livro": "Emitente (Livro)",
        "Emitente_XML": "Emitente (XML)",
        "CNPJ/CPF_Livro": "CNPJ/CPF (Livro)",
        "CNPJ/CPF_XML": "CNPJ/CPF (XML)",
        "CFOP_Livro": "CFOP (Livro)",
        "CFOP_Livro_Descrição": "Descrição CFOP (Livro)",
        "CFOP_XML": "CFOP (XML)",
        "CFOP_XML_Descrição": "Descrição CFOP (XML)",
        "Data_Emissao_Livro": "Data Emissão (Livro)",
        "Data_Emissao_XML": "Data Emissão (XML)",
        "VALOR_CONTÁBIL_Livro": "Valor Contábil (Livro)",
        "Valor_XML": "Valor (XML)",
        "BASE_CALCULO_Livro": "Base Cálculo (Livro)",
        "ISENTAS_OU_N_TRIB_Livro": "Isentas/Não Trib. (Livro)",
        "OUTRAS_Livro": "Outras (Livro)",
        "ICMS_Livro": "ICMS (Livro)",
        "Descrição_Produto_XML": "Descrição Produto (XML)",
        "Base_ICMS_XML_Total": "Base ICMS Total (XML)",
        "Valor_ICMS_XML_Total": "Valor ICMS Total (XML)",
        "Correção": "Correção Sugerida",
        "Status": "Status"
    }
    # Ordem de colunas para melhor análise
    ordem_visualizacao = [
    'Valor_vST', 'Alerta_ICMS_ST',
        "Número_Livro", "Número_XML",
        "Série_Livro", "Série_XML",
        "Espécie_Livro", "Espécie_XML",
        "Emitente_Livro", "Emitente_XML",
        "CNPJ/CPF_Livro", "CNPJ/CPF_XML",
        "CFOP_Livro", "CFOP_Livro_Descrição",
        "CFOP_XML", "CFOP_XML_Descrição",
        "Data_Emissao_Livro", "Data_Emissao_XML",
        "VALOR_CONTÁBIL_Livro", "Valor_XML",
        "BASE_CALCULO_Livro", "ISENTAS_OU_N_TRIB_Livro", "OUTRAS_Livro",
        "ICMS_Livro", "Descrição_Produto_XML",
        "Base_ICMS_XML_Total", "Valor_ICMS_XML_Total",
        "Correção", "Status"
    ]
    # Filtra colunas existentes e aplica renomeação
    colunas_viz = [
    "Número_Livro",
    "Número_XML",
    "Série_Livro",
    "Série_XML",
    "Espécie_Livro",
    "Espécie_XML",
    "Emitente_Livro",
    "Emitente_XML",
    "CNPJ/CPF_Livro",
    "CNPJ/CPF_XML",
    "CFOP_Livro",
    "CFOP_Livro_Descrição",
    "CFOP_XML",
    "CFOP_XML_Descrição",
    "Data_Emissao_Livro",
    "Data_Emissao_XML",
    "VALOR_CONTÁBIL_Livro",
    "Valor_XML",
    "ICMS_Livro",
    "Valor_ICMS_XML_Total",
    "BASE_CALCULO_Livro",
    "ISENTAS_OU_N_TRIB_Livro",
    "OUTRAS_Livro",
    "Descrição_Produto_XML",
    "Base_ICMS_XML_Total",
    "Valor_vST",
    "Alerta_ICMS_ST",
    "Status",
    "Correção"
]
    df_viz = df_final_renomeado[colunas_viz].rename(columns=display_names)
    
    
    tab_resultado, tab_auditoria, tab_livro = st.tabs(["📋 Resultado Geral", "🧪 Auditoria Detalhada", "📘 Livro Fiscal Original"])

    with tab_livro:
        st.markdown("### 📘 Visualização do Livro Fiscal Original")

        # Remover colunas indesejadas e reorganizar
        colunas_para_mostrar = [col for col in df_livro_original.columns if "CNPJ" not in col.upper()]
        df_livro_formatado = df_livro_original[colunas_para_mostrar].copy()

        st.dataframe(df_livro_formatado, use_container_width=True)

        csv_livro = df_livro_formatado.to_csv(index=False, sep=";", encoding="utf-8-sig")
        st.download_button(
            label="⬇️ Baixar Livro Fiscal (CSV)",
            data=csv_livro,
            file_name="livro_fiscal_original.csv",
            mime="text/csv"
        )

    with tab_auditoria:
        df_erros = df_final_renomeado[df_final_renomeado["Status"] != "✅ Validado"].copy()

        st.markdown("## 🧪 Auditoria Detalhada", unsafe_allow_html=True)
        st.markdown("<p style='color: gray'>Selecione um documento com divergência para análise detalhada lado a lado.</p>", unsafe_allow_html=True)

        opcoes = df_erros["Número_XML"].astype(str) + " - " + df_erros["Emitente_XML"]
        selecionado = st.selectbox("Selecionar Documento", options=opcoes.tolist())

        if selecionado:
            numero = selecionado.split(" - ")[0].strip()
            doc = df_erros[df_erros["Número_XML"].astype(str) == numero].iloc[0]

            col1, col2 = st.columns(2)

            with col1:
                st.subheader("📄 XML")
                st.write(f"**Emitente:** {doc['Emitente_XML']}")
                st.write(f"**CNPJ:** {doc['CNPJ/CPF_XML']}")
                st.write(f"**Data:** {doc['Data_Emissao_XML']}")
                st.write(f"**CFOP:** {doc['CFOP_XML']} - {doc.get('CFOP_XML_Descrição', '')}")
                st.write(f"**Valor:** {doc['Valor_XML']}")
                st.write(f"**ICMS:** {doc.get('Valor_ICMS_XML_Total', '')}")
                st.write(f"**Produto:** {doc.get('Descrição_Produto_XML', '')}")

            with col2:
                st.subheader("📃 Livro Fiscal")
                st.write(f"**Emitente:** {doc['Emitente_Livro']}")
                st.write(f"**CNPJ:** {doc['CNPJ/CPF_Livro']}")
                st.write(f"**Data:** {doc['Data_Emissao_Livro']}")
                st.write(f"**CFOP:** {doc['CFOP_Livro']} - {doc.get('CFOP_Livro_Descrição', '')}")
                st.write(f"**Valor Contábil:** {doc['VALOR_CONTÁBIL_Livro']}")
                st.write(f"**ICMS:** {doc.get('ICMS_Livro', '')}")
                st.write(f"**Base Cálculo:** {doc.get('BASE_CALCULO_Livro', '')}")
                st.write(f"**Isentas:** {doc.get('ISENTAS_OU_N_TRIB_Livro', '')}")
                st.write(f"**Outras:** {doc.get('OUTRAS_Livro', '')}")

            st.markdown("<hr style='margin-top: 1em;'>", unsafe_allow_html=True)
            st.markdown(f"### ⚠️ Divergências Encontradas\n**Status:** {doc['Status']}")
            st.warning(doc['Correção'])

            anotacao = st.text_area("📌 Anotações do Auditor", key=f"note_{numero}")

            if st.button("✅ Marcar como Corrigido"):
                st.success("Documento marcado como corrigido visualmente (não afeta os dados).")


    
    with tab_resultado:


    
        # ── Filtros Avançados dentro da aba Resultado Geral


    
        with st.expander("🔍 Filtros Avançados"):


    
            colf1, colf2 = st.columns(2)


    
            with colf1:


    
                filtro_cfop = st.multiselect("CFOP (XML)", options=sorted(df_final_renomeado["CFOP_XML"].dropna().unique()))


    
                filtro_modelo = st.multiselect("Modelo", options=sorted(df_final_renomeado["Espécie_XML"].dropna().unique()))


    
                filtro_status = st.multiselect("Status", options=sorted(df_final_renomeado["Status"].dropna().unique()))


    
            with colf2:


    
                filtro_numero = st.text_input("Número do Documento")


    
                filtro_emitente = st.text_input("Emitente (XML)")


    
                data_ini = st.date_input("Data Inicial (XML)", value=None)


    
                data_fim = st.date_input("Data Final (XML)", value=None)


    
    


    
            df_viz = df_viz[


    
                df_viz["CFOP (XML)"].isin(filtro_cfop) if filtro_cfop else True &


    
                df_viz["Modelo (XML)"].isin(filtro_modelo) if filtro_modelo else True &


    
                df_viz["Status"].isin(filtro_status) if filtro_status else True &


    
                df_viz["Número (XML)"].str.contains(filtro_numero, case=False) if filtro_numero else True &


    
                df_viz["Emitente (XML)"].str.contains(filtro_emitente, case=False) if filtro_emitente else True


    
            ]

        
        def aplicar_estilo(df):
            def color_status(val):
                if "Validado" in val:
                    return "background-color: #d4edda; color: #155724;"
                elif "Divergência" in val:
                    return "background-color: #fff3cd; color: #856404;"
                elif "❌" in val or "💸" in val:
                    return "background-color: #f8d7da; color: #721c24;"
                return ""
            return df.style.applymap(color_status, subset=["Status"])
        st.dataframe(aplicar_estilo(df_viz), use_container_width=True, height=600)
        import unicodedata
        def remover_acentos_texto(texto):
            try:
                return unicodedata.normalize('NFKD', str(texto)).encode('ASCII', 'ignore').decode('ASCII')
            except:
                return texto

        df_export_csv = df_viz.copy()
        df_export_csv.columns = [remover_acentos_texto(c) for c in df_export_csv.columns]
        df_export_csv = df_export_csv.applymap(remover_acentos_texto)
        csv_buffer = df_export_csv.to_csv(index=False, sep=';', encoding='utf-8-sig')
        st.download_button(
            label='⬇️ Baixar Resultado (CSV)',
            data=csv_buffer,
            file_name='resultado_validado.csv',
            mime='text/csv'
        )

        
except Exception:
    st.warning("⚠️ Não foi possível carregar os CFOPs para extração de produtos e ICMS. Verifique o arquivo cfop_produto_icms.xlsx e seu formato.")
    cfops_entrada_validos = []

st.markdown("""
    <hr style='margin-top: 2rem;'>
    <div style='text-align: right; font-size: 0.85rem; color: #666;'>
        Francisi Agro - Versão 1.0 - Desenvolvido por Danilo Araújo
    </div>
""", unsafe_allow_html=True)


# Processa XMLs se houver pelo menos um carregado
if uploaded_files:
    from io import BytesIO
    import xml.etree.ElementTree as ET

    xml_dfs = []
    for uploaded_file in uploaded_files:
        uploaded_file.seek(0)
        conteudo = uploaded_file.read()

        if not conteudo.strip():
            st.warning(f"O arquivo '{uploaded_file.name}' está vazio.")
            continue

        try:
            tree = ET.parse(BytesIO(conteudo))
            root = tree.getroot()
            df = extrair_dados_xml(root)  # <- ajuste conforme sua função real
            xml_dfs.append(df)
        except ET.ParseError:
            st.error(f"Erro ao processar o XML '{uploaded_file.name}': arquivo malformado.")

    df_xml_total = pd.concat(xml_dfs, ignore_index=True)




def gerar_danfe_html(xml_bytes):
    import xml.etree.ElementTree as ET
    import datetime

    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

    def tag(el, name):
        return el.findtext(f"nfe:{name}", default="", namespaces=ns) if el is not None else ""

    def formatar_data_br(data_iso):
        try:
            return datetime.datetime.fromisoformat(data_iso.replace("Z", "")).strftime("%d/%m/%Y")
        except:
            return data_iso

    def formatar_reais(valor):
        try:
            v = float(valor)
            return f"R$ {v:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
        except:
            return valor

    root = ET.fromstring(xml_bytes)
    ide = root.find(".//nfe:ide", ns)
    emit = root.find(".//nfe:emit", ns)
    dest = root.find(".//nfe:dest", ns)
    total = root.find(".//nfe:total/nfe:ICMSTot", ns)
    produtos = root.findall(".//nfe:det", ns)
    transp = root.find(".//nfe:transporta", ns)
    infCpl = root.find(".//nfe:infCpl", ns)
    chave = tag(root.find(".//nfe:infProt", ns), 'chNFe')

    html = f"""
<div style="font-family: monospace; border:1px solid #444; padding:20px;">
<h2 style="text-align:center;">DANFE - Documento Auxiliar da NF-e</h2>
<p style="text-align:center;">Chave de Acesso: <strong>{chave}</strong></p>
<hr>
<h3>Emitente</h3>
<p><strong>{tag(emit, 'xNome')}</strong><br>
CNPJ: {tag(emit, 'CNPJ')} | IE: {tag(emit, 'IE')}<br>
{tag(emit.find('nfe:enderEmit', ns), 'xLgr')}, {tag(emit.find('nfe:enderEmit', ns), 'nro')}<br>
{tag(emit.find('nfe:enderEmit', ns), 'xBairro')} - {tag(emit.find('nfe:enderEmit', ns), 'xMun')} - {tag(emit.find('nfe:enderEmit', ns), 'UF')} - CEP: {tag(emit.find('nfe:enderEmit', ns), 'CEP')}
</p>
<h3>Destinatário</h3>
<p><strong>{tag(dest, 'xNome')}</strong><br>
CPF/CNPJ: {tag(dest, 'CPF') or tag(dest, 'CNPJ')} | IE: {tag(dest, 'IE')}<br>
{tag(dest.find('nfe:enderDest', ns), 'xLgr')}, {tag(dest.find('nfe:enderDest', ns), 'nro')}<br>
{tag(dest.find('nfe:enderDest', ns), 'xBairro')} - {tag(dest.find('nfe:enderDest', ns), 'xMun')} - {tag(dest.find('nfe:enderDest', ns), 'UF')} - CEP: {tag(dest.find('nfe:enderDest', ns), 'CEP')}
</p>
<h3>Nota Fiscal</h3>
<p>Número: {tag(ide, 'nNF')} | Série: {tag(ide, 'serie')} | Modelo: {tag(ide, 'mod')}<br>
Emissão: {formatar_data_br(tag(ide, 'dhEmi'))} | Natureza: {tag(ide, 'natOp')}</p>
<h3>Produtos</h3>
<table border="1" width="100%" style="border-collapse: collapse; text-align:left; font-size: 13px;">
<tr>
<th>Item</th><th>Código</th><th>Descrição</th><th>CFOP</th><th>NCM</th><th>Qtd</th>
<th>Unitário</th><th>Total</th><th>CST/CSOSN</th><th>Alíquota</th>
<th>BC ICMS</th><th>Valor ICMS</th><th>BC ST</th><th>Valor ST</th>
<th>IPI</th><th>Frete</th><th>Outras Despesas</th><th>Seguro</th>
</tr>"""

    for det in produtos:
        prod = det.find("nfe:prod", ns)
        imposto = det.find("nfe:imposto", ns)
        # ICMS detail
        icms_root = imposto.find("nfe:ICMS", ns)
        icms_detail = list(icms_root)[0] if icms_root is not None and len(icms_root) else None
        # CST/CSOSN e Alíquota

        icms_root = imposto.find("nfe:ICMS", ns)
        icms_detail = list(icms_root)[0] if icms_root is not None and len(icms_root) else None
        # CST/CSOSN e Alíquota
        cst = ''
        aliquota = ''
        if icms_detail is not None:
            # Simples Nacional regime: blank aliquota
            local_name = icms_detail.tag.split('}')[-1]
            if local_name.startswith('ICMSSN'):
                cst = tag(icms_detail, 'CSOSN')
                aliquota = ''
            else:
                cst = tag(icms_detail, 'CST') or tag(icms_detail, 'CSOSN')
                aliquota_raw = tag(icms_detail, 'pICMS') or tag(icms_detail, 'pICMSST')
                try:
                    value = float(aliquota_raw)
                    aliquota = f"{value:.2f}".rstrip('0').rstrip('.') + '%'
                except:
                    aliquota = aliquota_raw + '%' if aliquota_raw else ''
# BC ICMS e Valor ICMS
        bc_icms = formatar_reais(tag(icms_detail, 'vBC')) if icms_detail is not None else ''
        valor_icms = formatar_reais(tag(icms_detail, 'vICMS')) if icms_detail is not None else ''
        bc_st = formatar_reais(tag(icms_detail, 'vBCST')) if icms_detail is not None else ''
        valor_st = formatar_reais(tag(icms_detail, 'vICMSST')) if icms_detail is not None else ''
        
        ipi_tag = det.find('nfe:imposto/nfe:IPI/nfe:IPITrib', ns)
        if ipi_tag is not None:
            ipi_val = formatar_reais(tag(ipi_tag, 'vIPI'))
        else:
            ipint = det.find('nfe:imposto/nfe:IPI/nfe:IPINT', ns)
            ipi_val = tag(ipint, 'CST') if ipint is not None else ''
        # Frete, seguro, outras despesas (totais)
        frete_val = formatar_reais(tag(total, 'vFrete'))
        seg_val = formatar_reais(tag(total, 'vSeg'))
        outras_val = formatar_reais(tag(total, 'vOutro'))

        html += f"""
<tr>
<td>{det.attrib.get('nItem')}</td>
<td>{tag(prod, 'cProd')}</td><td>{tag(prod, 'xProd')}</td><td>{tag(prod, 'CFOP')}</td><td>{tag(prod, 'NCM')}</td><td>{tag(prod, 'qCom')}</td>
<td>{formatar_reais(tag(prod, 'vUnCom'))}</td><td>{formatar_reais(tag(prod, 'vProd'))}</td>
<td>{cst}</td><td>{aliquota}</td>
<td>{bc_icms}</td><td>{valor_icms}</td><td>{bc_st}</td><td>{valor_st}</td>
<td>{ipi_val}</td><td>{frete_val}</td><td>{outras_val}</td><td>{seg_val}</td>
</tr>"""

    html += "</table>"
    html += f"""
<h3>Totais</h3>
<p>Base ICMS: {formatar_reais(tag(total, 'vBC'))}<br>
Valor ICMS: {formatar_reais(tag(total, 'vICMS'))}<br>
Valor ST: {formatar_reais(tag(total, 'vST'))}<br>
Valor Produtos: {formatar_reais(tag(total, 'vProd'))}<br>
IPI: {formatar_reais(tag(total, 'vIPI'))}<br>
Frete: {formatar_reais(tag(total, 'vFrete'))}<br>
Outras Despesas: {formatar_reais(tag(total, 'vOutro'))}<br>
Seguro: {formatar_reais(tag(total, 'vSeg'))}<br>
Valor NF: <strong>{formatar_reais(tag(total, 'vNF'))}</strong></p>"""

    if transp is not None:
        html += f"""
<h3>Transportadora</h3>
<p>{tag(transp, 'xNome')} - CNPJ: {tag(transp, 'CNPJ')}<br>
Placa: {tag(transp, 'placa')} - UF: {tag(transp, 'UF')}</p>"""

    if infCpl is not None and infCpl.text:
        html += f"""
<h3>Informações adicionais</h3>
<p>{infCpl.text}</p>"""

    html += "</div>"
    return html




# 📄 Aba de visualização de DANFE por XML
if st.sidebar.checkbox("📄 XML - Visualizar DANFE"):
    st.header("📄 Visualizador de DANFE via XML")
    arquivos = st.file_uploader("Selecione arquivos XML", type="xml", accept_multiple_files=True)
    filtro_numero   = st.text_input("🔎 Filtrar por Número da NF-e",   key="filter_num_danfe")
    filtro_emitente = st.text_input("🔎 Filtrar por Emitente da NF-e", key="filter_emit_danfe")
    if arquivos:
        # Aplica filtros de número e emitente
        arquivos_filtrados = []
        for fxml in arquivos:
            fxml.seek(0)
            xml_bytes = fxml.read()
            try:
                root = ET.fromstring(xml_bytes)
                df_dados = extrair_dados_xml(root)
                if df_dados.empty:
                    continue  # pula arquivos não-NF-e
                numero = df_dados["Número_XML"].iloc[0]
                emitente = df_dados["Emitente_XML"].iloc[0]
            except Exception:
                arquivos_filtrados.append(fxml)
                continue
            fxml.seek(0)
            if filtro_numero and filtro_numero not in numero:
                continue
            if filtro_emitente and filtro_emitente.lower() not in emitente.lower():
                continue
            arquivos_filtrados.append(fxml)
        arquivos = arquivos_filtrados
        for i, file in enumerate(arquivos):
            st.markdown(f"### 📎 {file.name}")
            if st.button("🔍 Visualizar DANFE", key=f"btn_xml_{i}"):
                try:
                    xml = file.read()
                    danfe = gerar_danfe_html(xml)
                    components.html(danfe, height=750, scrolling=True)
                except Exception as e:
                    st.error(f"Erro ao processar {file.name}: {e}")

import io
import zipfile
import os


def gerar_zip_xmls_livro():
    # Seleciona somente os arquivos XML válidos (que existem no Livro)
    valid_names = df_final_renomeado.loc[
        df_final_renomeado["Status"] != "❌ XML não encontrado no Livro",
        "Arquivo_XML"
    ].dropna().unique().tolist()
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as zf:
        for f in xml_files:
            if f.name in valid_names:
                f.seek(0)

                content = f.read()
                zf.writestr(f.name, content)
    buffer.seek(0)
    return buffer

# Botão de download em ZIP (mantido)

if st.button("📦 Baixar XMLs do Livro"):
    zip_buffer = gerar_zip_xmls_livro()
    st.download_button(
        label="Download dos XMLs em ZIP",
        data=zip_buffer,
        file_name="xmls_do_livro.zip",
        mime="application/zip"
    )
