import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

st.set_page_config(page_title="NFS a Pagar", layout="wide")

COLUMNS = ["FORNECEDOR", "CNPJ", "NUMERO", "DATA", "VALOR"]

def detect_and_load_excel(file) -> pd.DataFrame:
    xl = pd.ExcelFile(file)
    sheet = xl.sheet_names[0]
    raw = xl.parse(sheet, header=None)
    # Detecta linha do cabeçalho
    header_idx = None
    for i in range(min(10, len(raw))):
        row = raw.iloc[i].astype(str).str.upper().tolist()
        if any('FORNECEDOR' in x for x in row) and any('VALOR' in x for x in row):
            header_idx = i
            break
    if header_idx is None:
        header_idx = 0
    df = xl.parse(sheet, header=header_idx)

    # Normaliza nomes de colunas
    rename_map = {}
    for col in df.columns:
        up = str(col).strip().upper()
        if 'FORNECEDOR' in up: rename_map[col] = 'FORNECEDOR'
        elif up in ('CNPJ','CPF','CNPJ/CPF','CNPJ / CPF'): rename_map[col] = 'CNPJ'
        elif up in ('N°','Nº','NUMERO','N° NF','Nº NF','NF','N','N.'): rename_map[col] = 'NUMERO'
        elif 'DATA' in up: rename_map[col] = 'DATA'
        elif 'VALOR' in up: rename_map[col] = 'VALOR'
    df = df.rename(columns=rename_map)

    # Mantém apenas colunas alvo
    keep = [c for c in COLUMNS if c in df.columns]
    df = df[keep].copy()

    # Remove possíveis rodapés/totalizadores
    def is_footer(row):
        return (
            pd.isna(row.get('FORNECEDOR')) and
            pd.isna(row.get('NUMERO')) and
            pd.isna(row.get('DATA')) and
            pd.notna(row.get('VALOR'))
        )
    if 'FORNECEDOR' in df.columns:
        df = df[~df.apply(is_footer, axis=1)]

    # Tipagem
    if 'DATA' in df.columns:
        df['DATA'] = pd.to_datetime(df['DATA'], errors='coerce').dt.date
    if 'VALOR' in df.columns:
        df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce')
    if 'NUMERO' in df.columns:
        df['NUMERO'] = df['NUMERO'].astype(str).str.replace('.0','', regex=False).str.strip()
        df.loc[df['NUMERO'].isin(['nan','None','NaT','']), 'NUMERO'] = ''
    if 'CNPJ' in df.columns:
        df['CNPJ'] = df['CNPJ'].astype(str).str.replace('.0','', regex=False).str.strip()
        df.loc[df['CNPJ'].isin(['nan','None','NaT','']), 'CNPJ'] = ''

    df = df.dropna(how='all')

    # Garante todas as colunas na ordem
    for c in COLUMNS:
        if c not in df.columns:
            if c in ('FORNECEDOR','CNPJ','NUMERO'):
                df[c] = ''
            elif c == 'DATA':
                df[c] = pd.NaT
            elif c == 'VALOR':
                df[c] = 0.0
    return df[COLUMNS].reset_index(drop=True)

@st.cache_data
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='NFS A PAGAR')
    return output.getvalue()

st.title("Cadastro simples — NFS a Pagar")

# Sidebar: upload
st.sidebar.subheader("Importar base (opcional)")
uploaded = st.sidebar.file_uploader("Carregar Excel existente", type=['xlsx','xls'], key="uploader")

# Inicializa base
if 'data' not in st.session_state:
    st.session_state['data'] = pd.DataFrame(columns=COLUMNS)

if uploaded is not None:
    try:
        st.session_state['data'] = detect_and_load_excel(uploaded)
        st.sidebar.success("Base importada com sucesso.")
    except Exception as e:
        st.sidebar.warning(f"Falha ao ler o arquivo: {e}")

df = st.session_state['data']

# Formulário novo lançamento (evita piscadas entre reruns)
with st.form("novo_lancamento", clear_on_submit=True):
    st.write("### Novo lançamento")
    col1, col2 = st.columns([2,1])
    with col1:
        fornecedor = st.text_input("Fornecedor *", key="fornecedor_input")
        cnpj = st.text_input("CNPJ/CPF", key="cnpj_input")
        numero = st.text_input("Número da NF", key="numero_input")
    with col2:
        data_nf = st.date_input("Data", value=date.today(), key="data_input")
        valor = st.number_input("Valor (R$)", min_value=0.0, step=0.01, format="%.2f", key="valor_input")
    submitted = st.form_submit_button("Adicionar", use_container_width=True)
    if submitted:
        if not fornecedor.strip():
            st.error("Fornecedor é obrigatório.")
        else:
            new_row = {
                "FORNECEDOR": fornecedor.strip(),
                "CNPJ": cnpj.strip(),
                "NUMERO": numero.strip(),
                "DATA": data_nf,
                "VALOR": float(valor) if pd.notna(valor) else 0.0
            }
            st.session_state['data'] = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            st.success("Lançamento adicionado.")

# Filtros
st.write("### Registros")
fcol1, fcol2, fcol3 = st.columns([2,1,1])
with fcol1:
    filtro_forn = st.text_input("Filtrar por Fornecedor", key="filtro_forn")
with fcol2:
    filtro_num = st.text_input("Filtrar por Nº NF", key="filtro_num")
with fcol3:
    min_val = st.number_input("Valor mínimo (R$)", min_value=0.0, value=0.0, step=0.01, key="filtro_min_val")

fdf = st.session_state['data'].copy()
if filtro_forn:
    fdf = fdf[fdf['FORNECEDOR'].str.contains(filtro_forn, case=False, na=False)]
if filtro_num:
    fdf = fdf[fdf['NUMERO'].str.contains(filtro_num, case=False, na=False)]
fdf = fdf[fdf['VALOR'].fillna(0) >= min_val]

st.dataframe(fdf, use_container_width=True)

# Exclusão por índice (da base completa)
colA, colB, colC = st.columns(3)
with colA:
    st.info("Para excluir, informe os índices da base (não do filtro).")
with colB:
    indices_txt = st.text_input("Índices para excluir (ex: 0,3,7)", key="idx_excluir")
with colC:
    if st.button("Confirmar exclusão", key="btn_excluir"):
        try:
            idx = [int(x.strip()) for x in indices_txt.split(',') if x.strip() != '']
            st.session_state['data'] = st.session_state['data'].drop(idx, errors='ignore').reset_index(drop=True)
            st.success("Registros excluídos (se existiam).")
        except Exception as e:
            st.error(f"Erro ao excluir: {e}")

st.markdown("---")
total = st.session_state['data']['VALOR'].fillna(0).sum()
st.metric("Total a pagar (R$)", f"{total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

# Download SEM condicional (evita bug do DOM)
col1, col2 = st.columns(2)
with col1:
    xls = to_excel_bytes(st.session_state['data'])
    st.download_button(
        label="Baixar Excel atualizado",
        data=xls,
        file_name="nfs_a_pagar_atualizado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_xls"
    )
with col2:
    if st.button("Limpar tudo", key="clear_all"):
        st.session_state['data'] = pd.DataFrame(columns=COLUMNS)
        st.success("Base zerada.")
