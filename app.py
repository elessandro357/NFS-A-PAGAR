import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

# -----------------------------
# Config
# -----------------------------
st.set_page_config(page_title="NFS a Pagar", layout="wide")
COLUMNS = ["FORNECEDOR", "CNPJ", "NUMERO", "DATA", "VALOR"]

# -----------------------------
# Utils
# -----------------------------
def empty_df() -> pd.DataFrame:
    df = pd.DataFrame({c: [] for c in COLUMNS})
    # Tipos base
    df["FORNECEDOR"] = df["FORNECEDOR"].astype(object)
    df["CNPJ"] = df["CNPJ"].astype(object)
    df["NUMERO"] = df["NUMERO"].astype(object)
    df["DATA"] = pd.Series([], dtype="datetime64[ns]")
    df["VALOR"] = pd.Series([], dtype="float")
    return df

def detect_and_load_excel(file) -> pd.DataFrame:
    xl = pd.ExcelFile(file)
    sheet = xl.sheet_names[0]
    raw = xl.parse(sheet, header=None)

    # Detecta linha do cabeçalho
    header_idx = None
    for i in range(min(10, len(raw))):
        row = raw.iloc[i].astype(str).str.upper().tolist()
        if any("FORNECEDOR" in x for x in row) and any("VALOR" in x for x in row):
            header_idx = i
            break
    if header_idx is None:
        header_idx = 0

    df = xl.parse(sheet, header=header_idx)

    # Normaliza nomes
    rename_map = {}
    for col in df.columns:
        up = str(col).strip().upper()
        if "FORNECEDOR" in up: rename_map[col] = "FORNECEDOR"
        elif up in ("CNPJ","CPF","CNPJ/CPF","CNPJ / CPF"): rename_map[col] = "CNPJ"
        elif up in ("N°","Nº","NUMERO","N° NF","Nº NF","NF","N","N."): rename_map[col] = "NUMERO"
        elif "DATA" in up: rename_map[col] = "DATA"
        elif "VALOR" in up: rename_map[col] = "VALOR"
    df = df.rename(columns=rename_map)

    # Mantém e ordena colunas
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = None
    df = df[COLUMNS].copy()

    # Remove rodapés (totalizadores vazios com VALOR preenchido)
    def is_footer(row):
        return (
            pd.isna(row.get("FORNECEDOR"))
            and pd.isna(row.get("NUMERO"))
            and pd.isna(row.get("DATA"))
            and pd.notna(row.get("VALOR"))
        )
    df = df[~df.apply(is_footer, axis=1)]

    # Tipagem
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce")
    df["NUMERO"] = df["NUMERO"].astype(str).str.replace(".0","", regex=False).str.strip()
    df.loc[df["NUMERO"].isin(["nan","None","NaT",""]), "NUMERO"] = ""
    df["CNPJ"] = df["CNPJ"].astype(str).str.replace(".0","", regex=False).str.strip()
    df.loc[df["CNPJ"].isin(["nan","None","NaT",""]), "CNPJ"] = ""

    return df.reset_index(drop=True)

@st.cache_data
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    # Converte DATA para date (evita timezone/horário)
    out = df.copy()
    if "DATA" in out.columns:
        out["DATA"] = pd.to_datetime(out["DATA"], errors="coerce").dt.date
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="NFS A PAGAR")
    return buf.getvalue()

def brl(x: float) -> str:
    try:
        return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"

def fmt_date_col(series: pd.Series) -> pd.Series:
    s = pd.to_datetime(series, errors="coerce")
    return s.dt.strftime("%d/%m/%Y")

# -----------------------------
# Estado inicial
# -----------------------------
if "data" not in st.session_state:
    st.session_state["data"] = empty_df()

# -----------------------------
# UI
# -----------------------------
st.title("NFS a Pagar — simples e direto")

with st.sidebar:
    st.subheader("Importar base (opcional)")
    up = st.file_uploader("Carregar Excel", type=["xlsx","xls"], key="uploader_main")
    if up is not None:
        try:
            st.session_state["data"] = detect_and_load_excel(up)
            st.success("Base importada.")
        except Exception as e:
            st.error(f"Falha ao ler: {e}")

# Editor principal (CRUD completo)
st.markdown("#### Editar lançamentos")
st.caption("Use a grade abaixo para **inserir, editar ou remover linhas**. Depois clique em **Salvar alterações**.")
edited_df = st.data_editor(
    st.session_state["data"],
    key="grid_main",
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "FORNECEDOR": st.column_config.TextColumn("Fornecedor", required=True),
        "CNPJ": st.column_config.TextColumn("CNPJ/CPF"),
        "NUMERO": st.column_config.TextColumn("Número NF"),
        "DATA": st.column_config.DateColumn("Data", format="DD/MM/YYYY", step=1),
        "VALOR": st.column_config.NumberColumn("Valor (R$)", min_value=0.0, step=0.01, help="Informe em reais"),
    }
)

col_save, col_clear = st.columns([1,1])
with col_save:
    if st.button("💾 Salvar alterações", key="btn_save"):
        # Normaliza tipos ao salvar
        edited_df["DATA"] = pd.to_datetime(edited_df["DATA"], errors="coerce")
        edited_df["VALOR"] = pd.to_numeric(edited_df["VALOR"], errors="coerce").fillna(0.0)
        st.session_state["data"] = edited_df.reset_index(drop=True)
        st.success("Alterações salvas.")

with col_clear:
    if st.button("🗑️ Limpar tudo", key="btn_clear"):
        st.session_state["data"] = empty_df()
        st.success("Base zerada.")

st.markdown("---")

# Métrica e download SEM condicional (estável)
total = pd.to_numeric(st.session_state["data"]["VALOR"], errors="coerce").fillna(0.0).sum()
st.metric("Total a pagar", brl(total))

xls_bytes = to_excel_bytes(st.session_state["data"])
st.download_button(
    label="⬇️ Baixar Excel atualizado",
    data=xls_bytes,
    file_name="nfs_a_pagar_atualizado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="download_main",
    use_container_width=True
)

# Filtros de visualização (somente leitura)
with st.expander("Filtrar visualização (opcional)"):
    f1, f2, f3 = st.columns([2,1,1])
    with f1:
        filtro_forn = st.text_input("Filtrar por Fornecedor", key="filtro_forn_main")
    with f2:
        filtro_num = st.text_input("Filtrar por Nº NF", key="filtro_num_main")
    with f3:
        min_val = st.number_input("Valor mínimo (R$)", min_value=0.0, value=0.0, step=0.01, key="filtro_min_val_main")

    view = st.session_state["data"].copy()
    if filtro_forn:
        view = view[view["FORNECEDOR"].astype(str).str.contains(filtro_forn, case=False, na=False)]
    if filtro_num:
        view = view[view["NUMERO"].astype(str).str.contains(filtro_num, case=False, na=False)]
    view = view[pd.to_numeric(view["VALOR"], errors="coerce").fillna(0.0) >= min_val]

    if not view.empty:
        show = view.copy()
        show["DATA"] = fmt_date_col(show["DATA"])
        show["VALOR"] = pd.to_numeric(show["VALOR"], errors="coerce").fillna(0.0).map(brl)
        st.dataframe(show, use_container_width=True)
    else:
        st.info("Sem resultados para os filtros atuais.")
