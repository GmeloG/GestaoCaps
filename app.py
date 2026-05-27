import streamlit as st
import pandas as pd
import sqlite3
import io
import os
import sys
from datetime import datetime

# ── Configuration ─────────────────────────────────────────────────────────────
# When running as a PyInstaller .exe, use the exe's directory so the DB stays
# in the shared folder and not in the temporary extraction folder.
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(APP_DIR, "capsulas.db")

ESTADO_OPTIONS = ["Novo", "Usado", "Amostras", "Stock", "S/ID", "NOK", "Outros"]

ESTADO_COLORS = {
    "Novo":     "#d4edda",
    "Usado":    "#fff3cd",
    "Amostras": "#cce5ff",
    "NOK":      "#f8d7da",
    "Stock":    "#e2e3e5",
    "S/ID":     "#fce8b2",
    "Outros":   "#e8d5f5",
}

SORT_MAP = {
    "Localização":  "localizacao",
    "Projeto":      "projeto",
    "Estado":       "estado",
    "Código":       "codigo",
    "Data receção": "data_rececao",
    "Quantidade":   "qtd",
}


# ── Database ───────────────────────────────────────────────────────────────────
def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db():
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS capsulas (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                localizacao  TEXT,
                projeto      TEXT,
                codigo       TEXT,
                data_rececao TEXT,
                qtd          TEXT,
                estado       TEXT,
                obs          TEXT,
                source       TEXT DEFAULT 'excel',
                created_at   TEXT DEFAULT (datetime('now')),
                updated_at   TEXT DEFAULT (datetime('now'))
            )
        """)
        # Migration: add source column to existing databases
        try:
            conn.execute("ALTER TABLE capsulas ADD COLUMN source TEXT DEFAULT 'excel'")
        except Exception:
            pass
        conn.commit()


def db_is_empty() -> bool:
    with get_conn() as conn:
        return conn.execute("SELECT COUNT(*) FROM capsulas").fetchone()[0] == 0


def load_all() -> pd.DataFrame:
    with get_conn() as conn:
        return pd.read_sql("SELECT * FROM capsulas ORDER BY id", conn)


def insert_row(d: dict):
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO capsulas (localizacao,projeto,codigo,data_rececao,qtd,estado,obs,source) "
            "VALUES (?,?,?,?,?,?,?,'manual')",
            (d["localizacao"], d["projeto"], d["codigo"],
             d["data_rececao"], d["qtd"], d["estado"], d["obs"]),
        )
        conn.commit()


def update_row(row_id: int, d: dict):
    with get_conn() as conn:
        conn.execute(
            "UPDATE capsulas SET localizacao=?,projeto=?,codigo=?,data_rececao=?,qtd=?,"
            "estado=?,obs=?,updated_at=datetime('now') WHERE id=?",
            (d["localizacao"], d["projeto"], d["codigo"],
             d["data_rececao"], d["qtd"], d["estado"], d["obs"], row_id),
        )
        conn.commit()


def delete_row(row_id: int):
    with get_conn() as conn:
        conn.execute("DELETE FROM capsulas WHERE id=?", (row_id,))
        conn.commit()


def bulk_import(df: pd.DataFrame):
    with get_conn() as conn:
        # Remove only Excel-sourced records, keep manual ones
        conn.execute("DELETE FROM capsulas WHERE source='excel' OR source IS NULL")
        # Reset auto-increment to max manual ID so Excel records restart from 1
        # (sqlite_sequence has no PK, so INSERT OR REPLACE creates duplicates — use DELETE+INSERT)
        max_manual = conn.execute(
            "SELECT COALESCE(MAX(id), 0) FROM capsulas WHERE source='manual'"
        ).fetchone()[0]
        conn.execute("DELETE FROM sqlite_sequence WHERE name='capsulas'")
        conn.execute(
            "INSERT INTO sqlite_sequence (name, seq) VALUES ('capsulas', ?)",
            (max_manual,),
        )
        for _, r in df.iterrows():
            conn.execute(
                "INSERT INTO capsulas (localizacao,projeto,codigo,data_rececao,qtd,estado,obs,source) "
                "VALUES (?,?,?,?,?,?,?,'excel')",
                (_safe(r, "localizacao"), _safe(r, "projeto"), _safe(r, "codigo"),
                 _safe(r, "data_rececao"), _safe(r, "qtd"), _safe(r, "estado"), _safe(r, "obs")),
            )
        conn.commit()


def _safe(row, col):
    v = row.get(col, "")
    return "" if pd.isna(v) or str(v) in ("nan", "NaT", "None") else str(v).strip()


# ── Excel helpers ──────────────────────────────────────────────────────────────
def _normalize_estado(v) -> str:
    if pd.isna(v) or str(v).strip() in ("", "nan"):
        return ""
    mapping = {
        "novo": "Novo", "novas": "Novo", "new": "Novo",
        "usado": "Usado", "usadas": "Usado", "usadas e novas": "Usado",
        "amostras": "Amostras", "amostra": "Amostras",
        "s/id": "S/ID", "sem id": "S/ID", "sem identificação": "S/ID",
        "stock": "Stock",
        "nok": "NOK",
        "outros": "Outros",
    }
    return mapping.get(str(v).lower().strip(), str(v).strip())


def _detect_col(columns: list) -> dict:
    targets = {
        "localizacao":  ["local", "loca"],
        "projeto":      ["projeto", "project"],
        "codigo":       ["digo", "cod", "code"],
        "data_rececao": ["data", "rece"],
        "qtd":          ["qtd", "quant"],
        "estado":       ["estado", "state", "status"],
        "obs":          ["obs"],
    }
    result = {}
    for col in columns:
        col_l = col.lower()
        for internal, patterns in targets.items():
            if internal not in result and any(p in col_l for p in patterns):
                result[col] = internal
                break
    return result


def read_excel(file) -> pd.DataFrame:
    xl = pd.ExcelFile(file, engine="openpyxl")
    df = xl.parse(xl.sheet_names[0])
    df = df.rename(columns=_detect_col(df.columns.tolist()))
    if "estado" in df.columns:
        df["estado"] = df["estado"].apply(_normalize_estado)
    for col in ["localizacao", "projeto", "codigo", "data_rececao", "qtd", "estado", "obs"]:
        if col not in df.columns:
            df[col] = ""
    return df[["localizacao", "projeto", "codigo", "data_rececao", "qtd", "estado", "obs"]]


def find_excel_in_folder() -> str | None:
    """Return path of first .xlsx found in the app folder, or None."""
    for f in os.listdir(APP_DIR):
        if f.lower().endswith(".xlsx"):
            return os.path.join(APP_DIR, f)
    return None


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter

    out = io.BytesIO()
    export_df = df[["localizacao", "projeto", "codigo", "data_rececao", "qtd", "estado", "obs"]].copy()
    export_df.columns = ["Localização", "Projeto", "Código", "Data receção", "Qtd.", "Estado", "Obs."]
    export_df = export_df.replace("", pd.NA)

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="Cápsulas", index=False)
        ws = writer.sheets["Cápsulas"]

        hdr_fill = PatternFill("solid", fgColor="1F3864")
        hdr_font = Font(color="FFFFFF", bold=True, size=11)
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        color_map = {k: v.lstrip("#") for k, v in ESTADO_COLORS.items()}
        for row_idx in range(2, ws.max_row + 1):
            estado_val = str(ws.cell(row=row_idx, column=6).value or "")
            if estado_val in color_map:
                fill = PatternFill("solid", fgColor=color_map[estado_val])
                for col_idx in range(1, 8):
                    ws.cell(row=row_idx, column=col_idx).fill = fill

        for i, w in enumerate([12, 28, 16, 14, 6, 12, 45], start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[1].height = 20
        ws.freeze_panes = "A2"

    return out.getvalue()


# ── Streamlit app ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Gestão de Stock de Cápsulas",
    page_icon="🧩",
    layout="wide",
    initial_sidebar_state="collapsed",
)

init_db()

# Auto-load Excel on first run (empty DB)
if db_is_empty():
    excel_path = find_excel_in_folder()
    if excel_path:
        try:
            df_auto = read_excel(excel_path)
            bulk_import(df_auto)
        except Exception as e:
            pass

# CSS
st.markdown("""
<style>
.app-header {
    background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
    padding: 18px 24px;
    border-radius: 12px;
    margin-bottom: 24px;
    color: white;
}
.app-header h1 { margin: 0; font-size: 1.7rem; }
.app-header p  { margin: 4px 0 0; opacity: 0.85; font-size: 0.9rem; }

.stDataFrame { border-radius: 8px; overflow: hidden; }

.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: 0.78rem;
    font-weight: 600;
    margin: 2px;
    color: #333;
}

/* Statistics cards - VISIVEL */
.stat-card {
    background: #0052cc;
    border-radius: 12px;
    padding: 24px 16px;
    text-align: center;
    border: 3px solid #0066ff;
    box-shadow: 0 8px 20px rgba(0, 102, 255, 0.5);
}
.stat-card .stat-value {
    font-size: 2rem !important;
    font-weight: 800 !important;
    color: #ffffff !important;
    text-shadow: 0 2px 4px rgba(0, 0, 0, 0.5) !important;
    margin: 10px 0 8px 0 !important;
    line-height: 1 !important;
}
.stat-card .stat-label {
    font-size: 1rem !important;
    color: #ffffff !important;
    font-weight: 700 !important;
    margin: 10px 0 0 0 !important;
    text-shadow: 0 2px 4px rgba(0, 0, 0, 0.4) !important;
}
</style>
""", unsafe_allow_html=True)

# ── Session state ──────────────────────────────────────────────────────────────
for key, default in [
    ("confirm_delete", None),
    ("edit_id", None),
    ("show_form", False),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(
    '<div class="app-header">'
    '<h1>🧩 Gestão de Stock de Cápsulas</h1>'
    '<p>Visualize, edite e exporte o inventário de cápsulas</p>'
    '</div>',
    unsafe_allow_html=True,
)


# ══════════════════════════════════════════════════════════════════════════════
# ACTION BAR
# ══════════════════════════════════════════════════════════════════════════════
bar_cols = st.columns([1, 1, 1, 1])

# — Reload Excel
with bar_cols[0]:
    if st.button("🔄 Reload", use_container_width=True, type="primary"):
        excel_path = find_excel_in_folder()
        if excel_path:
            df_imp = read_excel(excel_path)
            bulk_import(df_imp)
            st.success(f"✅ Excel recarregado")
            st.rerun()

# — Add
with bar_cols[1]:
    if st.button("➕ Nova Cápsula", use_container_width=True):
        st.session_state.show_form = not st.session_state.show_form
        st.session_state.edit_id = None
        st.rerun()

# — Export
with bar_cols[2]:
    df_all = load_all()
    if not df_all.empty:
        st.download_button(
            "📤 Exportar Excel",
            data=to_excel_bytes(df_all),
            file_name=f"capsulas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# ══════════════════════════════════════════════════════════════════════════════
# ADD / EDIT FORM
# ══════════════════════════════════════════════════════════════════════════════
def _form(title: str, defaults: dict = None):
    d = defaults or {}
    with st.form("cap_form", clear_on_submit=True):
        st.markdown(f"**{title}**")
        c1, c2, c3 = st.columns(3)
        localizacao  = c1.text_input("Localização *",  value=d.get("localizacao", ""))
        projeto      = c1.text_input("Projeto",         value=d.get("projeto", ""))
        codigo       = c2.text_input("Código",           value=d.get("codigo", ""))
        data_rececao = c2.text_input("Data receção",     value=d.get("data_rececao", ""))
        qtd          = c3.text_input("Quantidade",       value=d.get("qtd", ""))
        estado_idx   = ESTADO_OPTIONS.index(d["estado"]) if d.get("estado") in ESTADO_OPTIONS else 0
        estado       = c3.selectbox("Estado", ESTADO_OPTIONS, index=estado_idx)
        obs          = st.text_area("Observações", value=d.get("obs", ""), height=80)

        sc1, sc2 = st.columns([1, 5])
        saved    = sc1.form_submit_button("💾 Guardar", type="primary")
        canceled = sc2.form_submit_button("Cancelar")

        if saved:
            if not localizacao.strip():
                st.error("A Localização é obrigatória.")
                return
            row = dict(
                localizacao=localizacao.strip(), projeto=projeto.strip(),
                codigo=codigo.strip(), data_rececao=data_rececao.strip(),
                qtd=qtd.strip(), estado=estado, obs=obs.strip(),
            )
            if st.session_state.edit_id:
                update_row(st.session_state.edit_id, row)
                st.success("✅ Registo atualizado.")
            else:
                insert_row(row)
                st.success("✅ Cápsula adicionada.")
            st.session_state.show_form = False
            st.session_state.edit_id = None
            st.rerun()

        if canceled:
            st.session_state.show_form = False
            st.session_state.edit_id = None
            st.rerun()


if st.session_state.show_form and st.session_state.edit_id is None:
    with st.container(border=True):
        _form("Adicionar nova cápsula")

if st.session_state.edit_id:
    df_tmp = load_all()
    row_match = df_tmp[df_tmp["id"] == st.session_state.edit_id]
    if not row_match.empty:
        with st.container(border=True):
            _form(f"Editar registo ID {st.session_state.edit_id}",
                  defaults=row_match.iloc[0].to_dict())


# ══════════════════════════════════════════════════════════════════════════════
# FILTERS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
f1, f2, f3, f4 = st.columns([3, 2, 2, 2])

search  = f1.text_input("🔍 Pesquisar", placeholder="Localização, Projeto, Código, Obs…")
estados = f2.multiselect("Filtrar por Estado", ESTADO_OPTIONS)
sort_by = f3.selectbox("Ordenar por", list(SORT_MAP.keys()), index=0)
asc     = f4.radio("Ordem", ["↑ Crescente", "↓ Decrescente"], horizontal=True) == "↑ Crescente"


# ══════════════════════════════════════════════════════════════════════════════
# LOAD + FILTER DATA
# ══════════════════════════════════════════════════════════════════════════════
df = load_all()

if df.empty:
    st.info("📂 Sem dados. Coloque um ficheiro .xlsx na pasta do servidor e clique em **🔄 Recarregar Excel**.")
    st.stop()

if search:
    mask = (
        df["localizacao"].str.contains(search, case=False, na=False) |
        df["projeto"].str.contains(search, case=False, na=False) |
        df["codigo"].str.contains(search, case=False, na=False) |
        df["obs"].str.contains(search, case=False, na=False)
    )
    df = df[mask]

if estados:
    df = df[df["estado"].isin(estados)]

df = df.sort_values(SORT_MAP[sort_by], ascending=asc, na_position="last")


# ══════════════════════════════════════════════════════════════════════════════
# STATISTICS
# ══════════════════════════════════════════════════════════════════════════════
total_qtd = pd.to_numeric(df["qtd"], errors="coerce").sum()
stats = [
    ("Registos", len(df)),
    ("Total cápsulas", int(total_qtd) if pd.notna(total_qtd) else "—"),
    ("Localizações", df["localizacao"].nunique()),
    ("Projetos", df["projeto"].nunique()),
    ("Total na BD", len(load_all())),
]

cols = st.columns(len(stats))
for col, (label, value) in zip(cols, stats):
    col.markdown(
        f'<div class="stat-card">'
        f'<div class="stat-value">{value}</div>'
        f'<div class="stat-label">{label}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

with st.expander("Ver detalhe por Estado e Projeto"):
    ec1, ec2 = st.columns(2)
    with ec1:
        st.markdown("**Por Estado**")
        st.dataframe(
            df.groupby("estado", dropna=False)
              .agg(Registos=("id", "count"))
              .sort_values("Registos", ascending=False)
              .reset_index().rename(columns={"estado": "Estado"}),
            hide_index=True, width='stretch',
        )
    with ec2:
        st.markdown("**Top 15 Projetos**")
        st.dataframe(
            df.groupby("projeto", dropna=False)
              .agg(Registos=("id", "count"))
              .sort_values("Registos", ascending=False)
              .head(15).reset_index().rename(columns={"projeto": "Projeto"}),
            hide_index=True, width='stretch',
        )


# ══════════════════════════════════════════════════════════════════════════════
# TABLE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"### 📋 Inventário — {len(df)} registo(s)")

display_df = df[["id", "localizacao", "projeto", "codigo",
                  "data_rececao", "qtd", "estado", "obs"]].copy()
display_df.columns = ["ID", "Localização", "Projeto", "Código",
                       "Data receção", "Qtd.", "Estado", "Obs."]
display_df = display_df.fillna("").replace("nan", "")

st.dataframe(
    display_df,
    hide_index=True,
    width='stretch',
)


# ══════════════════════════════════════════════════════════════════════════════
# EDIT / DELETE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("### ✏️ Editar / Remover registo")

all_ids = load_all()["id"].tolist()

st.markdown("**ID do registo**")
ec1, ec2, ec3, _ = st.columns([2, 1, 1, 4])

with ec1:
    selected_id = st.number_input(
        "id", min_value=1, step=1, label_visibility="collapsed",
        value=int(all_ids[0]) if all_ids else 1,
    )

with ec2:
    if st.button("✏️ Editar", use_container_width=True):
        if selected_id in all_ids:
            st.session_state.edit_id = int(selected_id)
            st.session_state.show_form = False
            st.rerun()
        else:
            st.error(f"ID {selected_id} não existe.")

with ec3:
    if st.button("🗑️ Apagar", use_container_width=True):
        if int(selected_id) in all_ids:
            st.session_state.confirm_delete = int(selected_id)
        else:
            st.error(f"ID {selected_id} não existe.")

if st.session_state.confirm_delete:
    cid = st.session_state.confirm_delete
    match = load_all()
    match = match[match["id"] == cid]
    if not match.empty:
        r = match.iloc[0]
        st.warning(
            f"⚠️ Confirma a eliminação do registo **ID {cid}** — "
            f"*{r['localizacao']}* / *{r['projeto']}*?"
        )
        dc1, dc2 = st.columns([1, 1])
        if dc1.button("✅ Sim, eliminar", type="primary"):
            delete_row(cid)
            st.session_state.confirm_delete = None
            st.success(f"Registo {cid} eliminado.")
            st.rerun()
        if dc2.button("❌ Cancelar"):
            st.session_state.confirm_delete = None
            st.rerun()