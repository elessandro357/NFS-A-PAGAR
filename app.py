import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

# =========================
# Configuração
# =========================
st.set_page_config(page_title="NFS a Pagar", layout="wide")
COLUMNS = ["FORNECEDOR", "CNPJ", "NUMERO", "DATA", "VALOR"]

# =========================
# Utilidades
# =========================
def empty_df() -> pd.DataFrame:
    df = pd.DataFrame(columns=COLUMNS)
    df["FORNECEDOR"] = df["FORNECEDOR"].astype(object)
    df["CNPJ"] = df["CNPJ"].astype(object)
    df["NUMERO"] = df["NUMERO"].astype(object)
    df["DATA"] = pd.Series([], dtype="datetime64[ns]")
    df["VALOR"] = pd.Series([], dtype="float")
    return df

def only_digits(s: str) -> str:
    if s is None or pd.isna(s):
        return ""
    return "".join(ch for ch in str(s) if ch.isdigit())

def norm_text(s: str) -> str:
    if s is None or pd.isna(s):
        return ""
    return str(s).strip()

def norm_numero(s: str) -> str:
    s = str(s)
    if s.lower() in ("nan", "nat", "none"):
        return ""
    return s.replace(".0", "").strip()

def parse_val(x):
    # Tenta converter para float, aceitando vírgula como decimal
    if isinstance(x, str):
        x = x.replace(".", "").replace(",", ".")
    return pd.to_numeric(x, errors="coerce")

def detect_and_load_excel(file) -> pd.DataFrame:
    xl = pd.ExcelFile(file)
    sheet = xl.sheet_names[0]
    raw = xl.parse(sheet, header=None)

    # Tenta achar linha de cabeçalho
    header_idx = None
    for i in range(min(10, len(raw))):
        row = raw.iloc[i].astype(str).str.upper().tolist()
        if any("FORNECEDOR" in x for x in row) and any("VALOR" in x for x in row):
            header_idx = i
            break
    if header_idx is None:
        header_idx = 0

    df = xl.parse(sheet, header=header_idx)

    # Renomeia colunas para o padrão
    rename_map = {}
    for col in df.columns:
        up = str(col).strip().upper()
        if "FORNECEDOR" in up: rename_map[col] = "FORNECEDOR"
        elif up in ("CNPJ","CPF","CNPJ/CPF","CNPJ / CPF"): rename_map[col] = "CNPJ"
        elif up in ("N°","Nº","NUMERO","N° NF","Nº NF","NF","N","N."): rename_map[col] = "NUMERO"
        elif "DATA" in up: rename_map[col] = "DATA"
        elif "VALOR" in up: rename_map[col] = "VALOR"
    df = df.rename(columns=rename_map)

    # Garante colunas alvo
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = None
    df = df[COLUMNS].copy()

    # Remove rodapés/totalizadores (linha sem fornecedor/numero/data e com valor preenchido)
    def is_footer(row):
        return (
            pd.isna(row.get("FORNECEDOR")) and
            pd.isna(row.get("NUMERO")) and
            pd.isna(row.get("DATA")) and
            pd.notna(row.get("VALOR"))
        )
    df = df[~df.apply(is_footer, axis=1)]

    # Normalizações
    df["FORNECEDOR"] = df["FORNECEDOR"].map(norm_text)
    df["CNPJ"] = df["CNPJ"].map(only_digits)
    df["NUMERO"] = df["NUMERO"].map(norm_numero)
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df["VALOR"] = df["VALOR"].apply(parse_val)

    # Descarta linhas 100% vazias
    df = df.dropna(how="all")
    # Preenche NaN de texto
    for c in ["FORNECEDOR","CNPJ","NUMERO"]:
        df[c] = df[c].fillna("")
    # Valor NaN -> 0.0
    df["VALOR"] = df["VALOR"].fillna(0.0)

    return df.reset_index(drop=True)

def make_key(df: pd.DataFrame) -> pd.Series:
    """
    Chave para deduplicação:
    - Preferência: CNPJ + NUMERO + DATA
    - Se não houver CNPJ: FORNECEDOR + NUMERO + DATA + VALOR
    """
    cnpj = df["CNPJ"].fillna("").astype(str)
    numero = df["NUMERO"].fillna("").astype(str)
    data = pd.to_datetime(df["DATA"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    valor = pd.to_numeric(df["VALOR"], errors="coerce").fillna(0.0).round(2).astype(str)
    fornecedor = df["FORNECEDOR"].fillna("").astype(str).str.upper().str.strip()

    key_with_cnpj = (cnpj != "")
    key = pd.Series(index=df.index, dtype="object")
    key.loc[key_with_cnpj] = cnpj[key_with_cnpj] + "|" + numero[key_with_cnpj] + "|" + data[key_with_cnpj]
    key.loc[~key_with_cnpj] = fornecedor[~key_with_cnpj] + "|" + numero[~key_with_cnpj] + "|" + data[~key_with_cnpj] + "|" + valor[~key_with_cnpj]
    return key

def merge_import(base: pd.DataFrame, incoming: pd.DataFrame, mode: str):
    """
    mode:
      - 'replace' -> substitui toda a base pelos dados importados
      - 'append_nodedup' -> anexa apenas os que não existem, conforme chave make_key
    """
    if mode == "replace":
        new_base = incoming.copy()
        return new_base.reset_index(drop=True), len(incoming), 0

    # append sem duplicar
    base = base.copy()
    base_key = make_key(base) if not base.empty else pd.Series([], dtype="object")
    inc_key = make_key(incoming)
    existing = set(base_key.tolist())
    mask_new = ~inc_key.isin(existing)
    added = incoming[mask_new].copy()

    if added.empty:
        return base.reset_index(drop=True), 0, (len(incoming) - 0)

    out = pd.concat([base, added], ignore_index=True)
    return out.reset_index(drop=True), len(added), (len(incoming) - len(added))

@st.cache_data
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    out = df.copy()
    out["DATA"] = pd.to_datetime(out["DATA"], errors="coerce").dt.date
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="NFS A PAGAR")
    return buf.getvalue()

def brl(v: float) -> str:
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "R$ 0,00"

def fmt_date_col(series: pd.Series) -> pd.Series:
    s = pd.to_datetime(series, errors="coerce")
    return s.dt.strftime("%d/%m/%Y")

# =========================
# Estado inicial
# =========================
if "data" not in st.session_state:
    st.session_state["data"] = empty_df()

# =========================
# UI
# =========================
st.title("NFS a Pagar — Importa e Cadastra")

with st.sidebar:
    st.subheader("Importar planilha e cadastrar")
    with st.form("form_import", clear_on_submit=True):
        mode = st.radio(
            "Modo de importação",
            options=["Substituir base", "Anexar (sem duplicar)"],
            index=1,
            key="mode_import"
        )
        up = st.file_uploader("Escolher Excel (.xlsx/.xls)", type=["xlsx","xls"], key="uploader_main")
        submitted = st.form_submit_button("Importar e cadastrar")
    if submitted:
        if up is None:
            st.warning("Selecione um arquivo.")
        else:
            try:
                incoming = detect_and_load_excel(up)
                if mode.startswith("Substituir"):
                    st.session_state["data"], added, skipped = merge_import(st.session_state["data"], incoming, mode="replace")
                else:
                    st.session_state["data"], added, skipped = merge_import(st.session_state["data"], incoming, mode="append_nodedup")

                total = pd.to_numeric(st.session_state["data"]["VALOR"], errors="coerce").fillna(0.0).sum()
                st.success(f"Importação concluída: {added} novos registros adicionados, {skipped} ignorados por duplicidade. Total atual: {brl(total)}")
            except Exception as e:
                st.error(f"Falha ao importar: {e}")

st.markdown("#### Lançamentos cadastrados")
st.caption("Edite diretamente na grade; clique em **Salvar alterações** para gravar na sessão.")

# Editor (CRUD)
edited_df = st.data_editor(
    st.session_state["data"],
    key="grid_main",
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "FORNECEDOR": st.column_config.TextColumn("Fornecedor", required=True),
        "CNPJ": st.column_config.TextColumn("CNPJ/CPF", help="Apenas números; usaremos para deduplicar com Nº+Data."),
        "NUMERO": st.column_config.TextColumn("Número NF"),
        "DATA": st.column_config.DateColumn("Data", format="DD/MM/YYYY", step=1),
        "VALOR": st.column_config.NumberColumn("Valor (R$)", min_value=0.0, step=0.01),
    }
)

col_save, col_clear = st.columns([1,1])
with col_save:
    if st.button("💾 Salvar alterações", key="btn_save"):
        edited_df["FORNECEDOR"] = edited_df["FORNECEDOR"].map(norm_text)
        edited_df["CNPJ"] = edited_df["CNPJ"].map(only_digits)
        edited_df["NUMERO"] = edited_df["NUMERO"].map(norm_numero)
        edited_df["DATA"] = pd.to_datetime(edited_df["DATA"], errors="coerce")
        edited_df["VALOR"] = edited_df["VALOR"].apply(parse_val).fillna(0.0)
        st.session_state["data"] = edited_df.reset_index(drop=True)
        st.success("Alterações salvas.")

with col_clear:
    if st.button("🗑️ Limpar tudo", key="btn_clear"):
        st.session_state["data"] = empty_df()
        st.success("Base zerada.")

st.markdown("---")

# Métricas + download (sempre renderizado)
base = st.session_state["data"]
total = pd.to_numeric(base["VALOR"], errors="coerce").fillna(0.0).sum()
forn_count = base["FORNECEDOR"].astype(str).str.strip().replace("", pd.NA).dropna().nunique()
qtd = len(base)
colm1, colm2, colm3 = st.columns(3)
colm1.metric("Total a pagar", brl(total))
colm2.metric("Fornecedores únicos", forn_count)
colm3.metric("Registros", qtd)

xls = to_excel_bytes(base)
st.download_button(
    label="⬇️ Baixar Excel atualizado",
    data=xls,
    file_name="nfs_a_pagar_atualizado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="download_xls",
    use_container_width=True
)

# Visualização filtrável (somente leitura)
with st.expander("Filtrar visualização (opcional)"):
    f1, f2, f3 = st.columns([2,1,1])
    with f1:
        filtro_forn = st.text_input("Filtrar por Fornecedor", key="filtro_forn_main")
    with f2:
        filtro_num = st.text_input("Filtrar por Nº NF", key="filtro_num_main")
    with f3:
        min_val = st.number_input("Valor mínimo (R$)", min_value=0.0, value=0.0, step=0.01, key="filtro_min_val_main")

    view = base.copy()
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
