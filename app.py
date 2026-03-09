"""
AP Contactos — INA  |  Gestão de contactos da Administração Pública portuguesa
"""
import html
import io
import re
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="AP Contactos", page_icon="🏛️", layout="wide")

# ─────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────
SHEETS_IGNORAR = {"Menu", "Conselho Editorial", "Conselho Estratégico"}

SHEET_CONFIG = {
    "Secretaria-Geral Governo ": {"header_row": 2, "cols": ["sigla_entidade","designacao","sigla_ministerio","ministerio","tipo_entidade","orgao_direcao","nome_dirigente","email","contacto"]},
    "Secretarias-Gerais":        {"header_row": 1, "cols": ["sigla_entidade","designacao","sigla_ministerio","ministerio","tipo_entidade","orgao_direcao","nome_dirigente","email","contacto"]},
    "Direções-Gerais":           {"header_row": 1, "cols": ["sigla_entidade","designacao","sigla_ministerio","ministerio","tipo_entidade","orgao_direcao","nome_dirigente","email","contacto"]},
    "Autoridades":               {"header_row": 1, "cols": ["sigla_entidade","designacao","sigla_ministerio","ministerio","tipo_entidade","orgao_direcao","nome_dirigente","email"]},
    "Academias":                 {"header_row": 1, "cols": ["sigla_entidade","designacao","sigla_ministerio","ministerio","tipo_entidade","orgao_direcao","nome_dirigente","email"]},
    "Cátedras Nacionais":        {"header_row": 1, "cols": ["sigla_entidade","designacao","sigla_ministerio","ministerio","tipo_entidade","orgao_direcao","nome_dirigente","contacto","email"]},
    "Inspeções-Gerais":          {"header_row": 1, "cols": ["sigla_entidade","designacao","sigla_ministerio","ministerio","tipo_entidade","orgao_direcao","nome_dirigente","email"]},
    "Institutos-Agências":       {"header_row": 1, "cols": ["sigla_entidade","designacao","sigla_ministerio","ministerio","tipo_entidade","orgao_direcao","nome_dirigente","email"]},
    "Direções-Regionais":        {"header_row": 1, "cols": ["convite","sigla_entidade","designacao","sigla_ministerio","ministerio","tipo_entidade","orgao_direcao","nome_dirigente","email"]},
    "Empresas Públicas":         {"header_row": 1, "cols": ["convite","sigla_entidade","designacao","sigla_ministerio","ministerio","tipo_entidade","orgao_direcao","nome_dirigente","email"]},
    "Fundações":                 {"header_row": 1, "cols": ["convite","designacao","tipo_entidade","orgao_direcao","nome_dirigente","email"]},
    "Comissões":                 {"header_row": 1, "cols": ["convite","sigla_entidade","designacao","sigla_ministerio","ministerio","tipo_entidade","orgao_direcao","nome_dirigente","email"]},
}

MAPA_SIOE = {
    "Direção-geral":"Direções-Gerais","Inspeção-geral":"Inspeções-Gerais","Inspeção Regional":"Inspeções-Gerais",
    "Secretaria-geral":"Secretarias-Gerais","Instituto Público":"Institutos-Agências","Agência":"Institutos-Agências",
    "Centro de Formação Profissional":"Institutos-Agências","Fundo Autónomo":"Institutos-Agências",
    "Fundo da Segurança Social":"Institutos-Agências","Banco Central":"Institutos-Agências",
    "Entidade Administrativa Independente":"Autoridades","Órgão Independente":"Autoridades",
    "Força de Segurança":"Autoridades","Forças Armadas":"Autoridades","Direção Regional":"Direções-Regionais",
    "Entidade Pública Empresarial":"Empresas Públicas","Entidade Pública Empresarial Regional":"Empresas Públicas",
    "Empresa Municipal":"Empresas Públicas","Empresa Intermunicipal":"Empresas Públicas",
    "Entidade Empresarial Municipal":"Empresas Públicas","Entidade Empresarial Regional":"Empresas Públicas",
    "Sociedade Anónima":"Empresas Públicas","Sociedade por Quotas":"Empresas Públicas",
    "Cooperativa":"Empresas Públicas","Agrupamento Complementar de Empresas":"Empresas Públicas",
    "Associação":"Associações","Fundação":"Fundações",
    "Estrutura temporária - comissão":"Comissões","Estrutura temporária - estrutura de missão":"Comissões",
    "Estrutura temporária - grupo de projeto":"Comissões","Estrutura temporária - grupo de trabalho":"Comissões",
    "Órgão consultivo":"Comissões","Serviço de Apoio":"Comissões",
    "Município":"Municípios","Junta de Freguesia":"Freguesias",
    "Associação de Municípios de fins específicos":"Municípios","Comunidade intermunicipal":"Municípios",
    "Federação de Municípios":"Municípios","Área Metropolitana":"Municípios",
    "Associação de Freguesias":"Freguesias","Serviço Municipalizado e Intermunicipalizado":"Municípios",
    "Tribunal":"Tribunais","Gabinete":"Gabinetes","Gabinete Ministro":"Gabinetes",
    "Gabinete 1.º Ministro":"Gabinetes","Gabinete Secretário de Estado":"Gabinetes",
    "Gabinete Secretário Regional":"Gabinetes","Gabinete Presidente Regional":"Gabinetes",
    "Gabinete Vice-Presidente Regional":"Gabinetes","Gabinete do Representante da República":"Gabinetes",
    "Unidade Orgânica de Ensino e Investigação":"Ensino e Investigação",
    "Unidade Orgânica de Investigação":"Ensino e Investigação",
    "Estabelecimento de educação e ensino básico e secundário":"Ensino e Investigação",
    "Entidade Regional de Turismo":"Outros","Estrutura atípica":"Outros",
}

COLS_BASE = ["sigla_entidade","designacao","ministerio","tipo_entidade",
             "orgao_direcao","nome_dirigente","email","contacto","categoria","fonte","id"]


# ─────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────
def limpar(v):
    if not isinstance(v, str):
        v = "" if v is None or (isinstance(v, float) and str(v) == "nan") else str(v)
    return "".join(c for c in v if unicodedata.category(c) not in ("Cf","Cc") or c in ("\n","\t")).strip()


def email_ok(e):
    e = limpar(str(e)).lower()
    if not e or e in ("nan","none","n/d"):
        return ""
    return e if re.match(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$", e) else ""


def garantir_cols(df):
    for col in COLS_BASE:
        if col not in df.columns:
            df[col] = ""
    return df


# ─────────────────────────────────────────────────────
# CARREGAR BASE PRINCIPAL
# ─────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def carregar_dados(file_bytes: bytes) -> pd.DataFrame:
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    frames = []
    for sheet in xl.sheet_names:
        ss = sheet.strip()
        if ss in SHEETS_IGNORAR:
            continue
        cfg = SHEET_CONFIG.get(sheet) or SHEET_CONFIG.get(ss)
        if cfg is None:
            continue
        try:
            df = xl.parse(sheet, header=cfg["header_row"])
            expected = cfg["cols"]
            df = df.iloc[:, :len(expected)]
            while len(df.columns) < len(expected):
                df[f"_x{len(df.columns)}"] = ""
            df.columns = expected
            df["categoria"] = ss
            df["fonte"] = "base_principal"
            frames.append(df)
        except Exception:
            pass
    if not frames:
        return pd.DataFrame(columns=COLS_BASE)
    dfall = pd.concat(frames, ignore_index=True)
    dfall["email"]      = dfall["email"].apply(email_ok)
    dfall["designacao"] = dfall["designacao"].apply(limpar)
    for col in ["nome_dirigente","ministerio","sigla_entidade","orgao_direcao","contacto","tipo_entidade","categoria","fonte"]:
        if col in dfall.columns:
            dfall[col] = dfall[col].apply(limpar).replace({"nan":"","None":""})
    dfall = dfall[dfall["designacao"].str.len() > 2].reset_index(drop=True)
    dfall["id"] = dfall.index
    return garantir_cols(dfall)


# ─────────────────────────────────────────────────────
# IMPORTAR SIOE
# ─────────────────────────────────────────────────────
def importar_sioe(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes), header=0)
    df.columns = [str(c).strip() for c in df.columns]

    tipo_cols     = [c for c in df.columns if c.startswith("Contacto") and c.endswith("Tipo")]
    contacto_cols = [c for c in df.columns if c.startswith("Contacto") and c.endswith("Contacto")]
    nome_cols     = [c for c in df.columns if "Membro" in c and c.endswith("Nome")]
    cargo_cols    = [c for c in df.columns if "Membro" in c and c.endswith("Cargo")]
    resp_cols     = [c for c in df.columns if "Membro" in c and "Responsável" in c]

    # Email e telefone — primeira ocorrência por tipo
    email_s = pd.Series("", index=df.index)
    tel_s   = pd.Series("", index=df.index)
    for tc, cc in zip(tipo_cols, contacto_cols):
        t = df[tc].fillna("").astype(str).str.strip().str.lower()
        v = df[cc].fillna("").astype(str).str.strip()
        email_s = email_s.where(~((t == "email")    & (email_s == "")), v.str.lower())
        tel_s   = tel_s.where(~((t == "telefone")   & (tel_s   == "")), v)
    email_s = email_s.where(
        email_s.str.match(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$"), ""
    )

    # Dirigente — Responsável == Sim; fallback: primeiro membro
    nome_s  = pd.Series("", index=df.index)
    cargo_s = pd.Series("", index=df.index)
    for nc, cc2, rc in zip(nome_cols, cargo_cols, resp_cols):
        is_r = df[rc].fillna("").astype(str).str.strip().str.lower() == "sim"
        need = nome_s == ""
        nome_s  = nome_s.where(~(is_r & need), df[nc].fillna("").astype(str).str.strip())
        cargo_s = cargo_s.where(~(is_r & need), df[cc2].fillna("").astype(str).str.strip())
    if nome_cols:
        nd = nome_s == ""
        nome_s  = nome_s.where(~nd, df[nome_cols[0]].fillna("").astype(str).str.strip())
        cargo_s = cargo_s.where(~nd, df[cargo_cols[0]].fillna("").astype(str).str.strip())

    def gc(n):
        return df.get(n, pd.Series("", index=df.index)).fillna("").astype(str).str.strip()

    novo = pd.DataFrame({
        "designacao":     gc("Designação"),
        "sigla_entidade": gc("Sigla"),
        "ministerio":     gc("Ministério/Secretaria Regional"),
        "tipo_entidade":  gc("Tipo de Entidade"),
        "email":          email_s,
        "contacto":       tel_s,
        "nome_dirigente": nome_s,
        "orgao_direcao":  cargo_s,
        "categoria":      gc("Tipo de Entidade").map(MAPA_SIOE).fillna("Outros"),
        "fonte":          "SIOE",
    })
    return novo[novo["designacao"].str.len() > 2].copy()


# ─────────────────────────────────────────────────────
# FUNDIR COM BASE
# ─────────────────────────────────────────────────────
def fundir_com_base(df_novo: pd.DataFrame) -> tuple:
    """
    Adiciona df_novo ao st.session_state['df'].
    Evita duplicar emails já existentes na base.
    Devolve (n_adicionados, n_emails).
    """
    df_novo = garantir_cols(df_novo.copy())

    # Remover entradas cujo email já existe na base (excepto sem email)
    emails_base = set(st.session_state["df"]["email"].str.lower())
    emails_base.discard("")
    mask_novo = ~df_novo["email"].isin(emails_base) | (df_novo["email"] == "")
    df_novo = df_novo[mask_novo].copy()

    if df_novo.empty:
        return 0, 0

    id_max = int(st.session_state["df"]["id"].max()) + 1
    df_novo = df_novo.reset_index(drop=True)
    df_novo["id"] = df_novo.index + id_max

    n_antes = len(st.session_state["df"])
    st.session_state["df"] = pd.concat(
        [st.session_state["df"], df_novo[COLS_BASE]], ignore_index=True
    )
    n_emails = int((df_novo["email"].str.len() > 3).sum())
    return len(st.session_state["df"]) - n_antes, n_emails


# ─────────────────────────────────────────────────────
# EXPORTAR EXCEL
# ─────────────────────────────────────────────────────
def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    cols = ["categoria","sigla_entidade","designacao","ministerio","tipo_entidade",
            "orgao_direcao","nome_dirigente","email","contacto","fonte"]
    cols_ok = [c for c in cols if c in df.columns]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df[cols_ok].to_excel(writer, sheet_name="Todos", index=False)
        for cat in sorted(df["categoria"].dropna().unique()):
            df[df["categoria"] == cat][cols_ok].to_excel(
                writer, sheet_name=str(cat)[:31], index=False
            )
    return buf.getvalue()


# ─────────────────────────────────────────────────────
# INICIALIZAR SESSION STATE
# ─────────────────────────────────────────────────────
defaults = {
    "df":          None,   # DataFrame principal
    "sel":         {},     # {id: bool} selecções para copiar emails
    "_file_key":   None,   # key do ficheiro base carregado
    "_ext_key":    None,   # key do ficheiro externo
    "_ext_bytes":  None,   # bytes do ficheiro externo (persiste após rerun)
    "_ext_name":   None,   # nome do ficheiro externo
    "import_msg":  None,   # mensagem de resultado da última importação
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ─────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────
st.markdown("""
<style>
.metric-box {background:white;border-radius:10px;padding:14px 18px;
  border-left:4px solid #1a56db;margin-bottom:10px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.metric-box .n {font-size:1.9rem;font-weight:700;color:#1a56db}
.metric-box .l {font-size:.82rem;color:#666}
.email-box {background:#f8f9fa;border:1px solid #dee2e6;border-radius:8px;
  padding:12px 16px;font-family:monospace;font-size:.84rem;
  max-height:320px;overflow-y:auto;white-space:pre-wrap;word-break:break-all}
.import-result {background:#d4edda;border:1px solid #c3e6cb;border-radius:8px;
  padding:14px 18px;margin:12px 0}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────
# SIDEBAR — FICHEIRO BASE
# ─────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏛️ AP Contactos")
    st.markdown("---")
    st.markdown("### 📂 Ficheiro principal")

    uploaded = st.file_uploader(
        "Excel INA (.xlsx)", type=["xlsx"], label_visibility="collapsed"
    )

    if uploaded is not None:
        file_bytes = uploaded.getvalue()
        file_key   = f"{uploaded.name}_{len(file_bytes)}"
        if st.session_state["_file_key"] != file_key:
            # Novo ficheiro → limpar tudo
            st.session_state["_file_key"]  = file_key
            st.session_state["df"]         = None
            st.session_state["sel"]        = {}
            st.session_state["_ext_key"]   = None
            st.session_state["_ext_bytes"] = None
            st.session_state["_ext_name"]  = None
            st.session_state["import_msg"] = None
            st.cache_data.clear()

        if st.session_state["df"] is None:
            with st.spinner("A carregar base..."):
                st.session_state["df"] = carregar_dados(file_bytes)

    if st.session_state["df"] is not None:
        n_tot  = len(st.session_state["df"])
        n_mail = (st.session_state["df"]["email"].str.len() > 3).sum()
        st.success(f"✅ {n_tot:,} registos  |  {n_mail:,} emails")
    else:
        st.info("Carrega o ficheiro Excel para começar.")

    st.markdown("---")
    st.markdown("### 🔍 Filtros")


# ─────────────────────────────────────────────────────
# GUARD
# ─────────────────────────────────────────────────────
if st.session_state["df"] is None:
    st.title("🏛️ Contactos — AP Portuguesa")
    st.info("👆 Carrega o ficheiro Excel principal na barra lateral para começar.")
    st.stop()

df = st.session_state["df"]   # referência viva — actualiza automaticamente


# ─────────────────────────────────────────────────────
# FILTROS SIDEBAR
# ─────────────────────────────────────────────────────
with st.sidebar:
    pesquisa = st.text_input("🔎 Pesquisar", "")
    cats_sel = st.multiselect("Categoria", sorted(df["categoria"].dropna().unique()))
    _prefixos = ("Ministério", "Secretaria", "Presidência", "Vice-Presidência")
    min_disp = sorted(
        m for m in df["ministerio"].fillna("").unique()
        if m and any(m.startswith(p) for p in _prefixos)
    )
    min_sel  = st.multiselect("Ministério / Tutela", min_disp)
    so_email = st.checkbox("Só com email", value=False)
    st.markdown("---")
    if st.button("🗑️ Limpar & recarregar", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.cache_data.clear()
        st.rerun()

df_f = df.copy()
if pesquisa:
    p = pesquisa.lower()
    m = (df_f["designacao"].str.lower().str.contains(p, na=False)
       | df_f["sigla_entidade"].str.lower().str.contains(p, na=False)
       | df_f["ministerio"].str.lower().str.contains(p, na=False)
       | df_f["nome_dirigente"].str.lower().str.contains(p, na=False)
       | df_f["email"].str.lower().str.contains(p, na=False))
    df_f = df_f[m]
if cats_sel:
    df_f = df_f[df_f["categoria"].isin(cats_sel)]
if min_sel:
    df_f = df_f[df_f["ministerio"].isin(min_sel)]
if so_email:
    df_f = df_f[df_f["email"].str.len() > 3]


# ─────────────────────────────────────────────────────
# MÉTRICAS
# ─────────────────────────────────────────────────────
st.title("🏛️ Contactos — AP Portuguesa")

# Banner de resultado de importação (persiste entre reruns)
if st.session_state.get("import_msg"):
    msg = st.session_state["import_msg"]
    st.success(msg["texto"])
    if st.button("✖ Fechar", key="fechar_msg"):
        st.session_state["import_msg"] = None
        st.rerun()

ev = df_f[df_f["email"].str.len() > 3]
cols_m = st.columns(4)
for col, num, lbl in zip(
    cols_m,
    [f"{len(df_f):,}", f"{len(ev):,}", f"{df_f['designacao'].nunique():,}",
     f"{df_f['ministerio'].replace('', pd.NA).dropna().nunique():,}"],
    ["Registos filtrados", "Com email", "Entidades únicas", "Ministérios"],
):
    col.markdown(
        f'<div class="metric-box"><div class="n">{num}</div><div class="l">{lbl}</div></div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Tabela",
    "☑️ Selecionar & Copiar Emails",
    "➕ Importar Base Externa",
    "📥 Descarregar",
])


# ══════════════════════════════════════════════════════
# TAB 1 — TABELA
# ══════════════════════════════════════════════════════
with tab1:
    lbl_map = {
        "categoria":"Categoria","sigla_entidade":"Sigla","designacao":"Entidade",
        "ministerio":"Ministério","orgao_direcao":"Órgão/Cargo","nome_dirigente":"Dirigente",
        "email":"Email","contacto":"Contacto","fonte":"Fonte",
    }
    cols_v = [c for c in lbl_map if c in df_f.columns]
    st.dataframe(df_f[cols_v].rename(columns=lbl_map),
                 use_container_width=True, height=520, hide_index=True)
    st.caption(f"A mostrar {len(df_f):,} de {len(df):,} registos totais")

    # ── Exportação rápida de emails do filtro actual ──
    df_em_f = df_f[df_f["email"].str.len() > 3].drop_duplicates(subset=["email"])
    n_em_f  = len(df_em_f)
    if n_em_f > 0:
        st.markdown("---")
        st.markdown(f"**📧 Exportação rápida — {n_em_f:,} emails no filtro actual:**")
        eq1, eq2, eq3 = st.columns(3)
        with eq1:
            txt_bcc = "; ".join(sorted(df_em_f["email"].tolist()))
            st.download_button(
                f"⬇️ BCC Outlook ({n_em_f:,})",
                data=txt_bcc.encode("utf-8"),
                file_name=f"emails_BCC_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain", use_container_width=True, type="primary",
                help="Separados por ; — cola no campo Cco do Outlook",
            )
        with eq2:
            txt_linhas = "\n".join(sorted(df_em_f["email"].tolist()))
            st.download_button(
                f"⬇️ Um por linha ({n_em_f:,})",
                data=txt_linhas.encode("utf-8"),
                file_name=f"emails_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain", use_container_width=True,
                help="Um email por linha",
            )
        with eq3:
            csv_em = df_em_f[["categoria","designacao","nome_dirigente","email"]].to_csv(
                index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                f"⬇️ CSV com nomes ({n_em_f:,})",
                data=csv_em,
                file_name=f"emails_nomes_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv", use_container_width=True,
                help="CSV com categoria, entidade, nome e email",
            )
    else:
        if len(df_f) > 0:
            st.caption("Nenhum registo com email no filtro actual.")


# ══════════════════════════════════════════════════════
# TAB 2 — SELECIONAR & COPIAR EMAILS
# ══════════════════════════════════════════════════════
with tab2:
    st.markdown("### ☑️ Selecionar organismos e copiar emails")
    df_ce = df_f[df_f["email"].str.len() > 3].copy()

    if df_ce.empty:
        st.warning("Nenhum registo com email nos filtros actuais.")
    else:
        c1b, c2b, c3b = st.columns([1, 1, 2])
        with c1b:
            if st.button("✅ Todos", use_container_width=True):
                for i in df_ce["id"].tolist():
                    st.session_state["sel"][i] = True
                st.rerun()
        with c2b:
            if st.button("❌ Limpar selecção", use_container_width=True):
                st.session_state["sel"] = {}
                st.rerun()
        with c3b:
            cat_r = st.selectbox("Adicionar categoria:", ["—"] + sorted(df_ce["categoria"].unique()), key="cat_r")
            if cat_r != "—" and st.button(f"➕ Adicionar '{cat_r}'", use_container_width=True):
                for i in df_ce[df_ce["categoria"] == cat_r]["id"].tolist():
                    st.session_state["sel"][i] = True
                st.rerun()

        st.markdown("---")
        for cat in sorted(df_ce["categoria"].unique()):
            sub = df_ce[df_ce["categoria"] == cat]
            n_s = sum(1 for i in sub["id"] if st.session_state["sel"].get(i, False))
            with st.expander(f"**{cat}** — {len(sub)} com email  |  ✓ {n_s} seleccionados"):
                sc1, sc2 = st.columns(2)
                with sc1:
                    if st.button(f"✅ Todos ({cat})", key=f"sa_{cat}", use_container_width=True):
                        for i in sub["id"].tolist():
                            st.session_state["sel"][i] = True
                        st.rerun()
                with sc2:
                    if st.button(f"❌ Nenhum ({cat})", key=f"sn_{cat}", use_container_width=True):
                        for i in sub["id"].tolist():
                            st.session_state["sel"][i] = False
                        st.rerun()
                for _, row in sub.iterrows():
                    idx = row["id"]
                    lbl_chk = f"{row['designacao']}  —  `{row['email']}`"
                    if str(row.get("nome_dirigente", "")).strip():
                        lbl_chk += f"  ({row['nome_dirigente']})"
                    v_novo = st.checkbox(
                        lbl_chk,
                        value=st.session_state["sel"].get(idx, False),
                        key=f"chk_{idx}",
                    )
                    if v_novo != st.session_state["sel"].get(idx, False):
                        st.session_state["sel"][idx] = v_novo

        st.markdown("---")
        ids_sel = [i for i, v in st.session_state["sel"].items() if v]
        df_sel  = df_ce[df_ce["id"].isin(ids_sel)]
        st.markdown(f"### 📋 {len(df_sel)} seleccionados")

        if not df_sel.empty:
            cf1, cf2 = st.columns([2, 1])
            with cf1:
                fmt = st.radio("Formato", ["Um por linha", "Separados por ; (BCC)", "CSV"], horizontal=True)
            with cf2:
                inc_nome = st.checkbox("Incluir nome", value=False)
                dedup    = st.checkbox("Sem duplicados", value=True)

            df_exp = df_sel.drop_duplicates(subset=["email"]) if dedup else df_sel
            lista  = sorted(df_exp["email"].tolist())

            if inc_nome:
                pares = df_exp[["nome_dirigente", "email"]]
                if fmt == "Um por linha":
                    texto = "\n".join(f"{r['nome_dirigente']} <{r['email']}>" for _, r in pares.iterrows())
                elif fmt == "Separados por ; (BCC)":
                    texto = "; ".join(f"{r['nome_dirigente']} <{r['email']}>" for _, r in pares.iterrows())
                else:
                    texto = "Nome,Email\n" + "\n".join(f'"{r["nome_dirigente"]}",{r["email"]}' for _, r in pares.iterrows())
            else:
                if fmt == "Um por linha":
                    texto = "\n".join(lista)
                elif fmt == "Separados por ; (BCC)":
                    texto = "; ".join(lista)
                else:
                    texto = "Email\n" + "\n".join(lista)

            st.markdown(f"**{len(lista)} emails:**")
            st.markdown(f'<div class="email-box">{html.escape(texto)}</div>', unsafe_allow_html=True)
            st.download_button(
                "⬇️ Descarregar (.txt)", data=texto.encode("utf-8"),
                file_name=f"emails_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain", use_container_width=True,
            )
        else:
            st.info("Nenhum organismo seleccionado. Usa as checkboxes acima.")


# ══════════════════════════════════════════════════════
# TAB 3 — IMPORTAR BASE EXTERNA
# ══════════════════════════════════════════════════════
with tab3:
    st.markdown("### ➕ Importar base de dados externa")

    uploaded_ext = st.file_uploader(
        "Carregar Excel externo (.xlsx)", type=["xlsx"],
        key="uploader_ext", label_visibility="collapsed",
    )

    # ── FIX PRINCIPAL: guardar bytes em session_state imediatamente ──
    # O file_uploader reseta para None após st.rerun().
    # Guardamos os bytes com a key, e usamo-los mesmo que uploaded_ext seja None.
    if uploaded_ext is not None:
        ext_key = f"{uploaded_ext.name}_{uploaded_ext.size}"
        if st.session_state["_ext_key"] != ext_key:
            # Novo ficheiro externo
            st.session_state["_ext_key"]   = ext_key
            st.session_state["_ext_bytes"] = uploaded_ext.getvalue()
            st.session_state["_ext_name"]  = uploaded_ext.name

    # Usar bytes de session_state (persiste após rerun)
    ext_bytes = st.session_state["_ext_bytes"]
    ext_name  = st.session_state["_ext_name"] or ""

    if ext_bytes is None:
        # Nenhum ficheiro externo carregado ainda
        st.info(
            "Suporta dois modos:\n\n"
            "- 🏛️ **Exportação do SIOE** (`ExportResultadosPesquisa*.xlsx`) — importação automática\n"
            "- 📋 **Qualquer outro Excel** — mapeias as colunas manualmente"
        )
    else:
        # ── Detectar se é SIOE ──
        parece_sioe = "export" in ext_name.lower() and "pesquisa" in ext_name.lower()
        if not parece_sioe:
            try:
                _t = pd.read_excel(io.BytesIO(ext_bytes), header=0, nrows=1)
                _t.columns = [str(c).strip() for c in _t.columns]
                parece_sioe = "Designação" in _t.columns and "Código SIOE" in _t.columns
            except Exception:
                pass

        st.markdown(f"**Ficheiro:** `{ext_name}`")
        if st.button("🗑️ Remover ficheiro externo", key="rm_ext"):
            st.session_state["_ext_key"]   = None
            st.session_state["_ext_bytes"] = None
            st.session_state["_ext_name"]  = None
            st.rerun()

        st.markdown("---")

        # ────────────────────────────
        # MODO SIOE
        # ────────────────────────────
        if parece_sioe:
            st.success("✅ Ficheiro reconhecido como exportação do **SIOE**")

            try:
                n_lin = len(pd.read_excel(io.BytesIO(ext_bytes), header=0, usecols=[0]))
            except Exception:
                n_lin = "?"

            cm, ci = st.columns([1, 2])
            with cm:
                st.metric("Entidades no ficheiro SIOE",
                          f"{n_lin:,}" if isinstance(n_lin, int) else n_lin)
                btn_sioe = st.button("🚀 Importar do SIOE", type="primary", use_container_width=True)
            with ci:
                st.info(
                    "Extrai automaticamente: designação, sigla, ministério, tipo de entidade, "
                    "email, telefone e dirigente responsável.\n\n"
                    "Não duplica emails já existentes na base principal."
                )

            if btn_sioe:
                with st.spinner("A processar o SIOE… (pode demorar ~10 segundos)"):
                    try:
                        df_novo = importar_sioe(ext_bytes)
                        if df_novo.empty:
                            st.error("Nenhum registo válido encontrado no ficheiro.")
                        else:
                            n_add, n_em = fundir_com_base(df_novo)
                            if n_add == 0:
                                st.warning("Todos os registos do SIOE já existem na base (emails duplicados).")
                            else:
                                por_cat = df_novo["categoria"].value_counts().head(10).to_dict()
                                linhas_cat = "\n".join(f"• {cat}: {n:,}" for cat, n in por_cat.items())
                                st.session_state["import_msg"] = {
                                    "texto": (
                                        f"✅ SIOE importado com sucesso! "
                                        f"**{n_add:,}** registos adicionados, **{n_em:,}** com email. "
                                        f"Base agora tem **{len(st.session_state['df']):,}** registos no total."
                                    )
                                }
                                st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao importar SIOE: {e}")

        # ────────────────────────────
        # MODO MANUAL
        # ────────────────────────────
        st.markdown("---")
        with st.expander(
            "🔧 Mapeamento manual de colunas",
            expanded=not parece_sioe,
        ):
            try:
                xl_ext = pd.ExcelFile(io.BytesIO(ext_bytes))
                sheets = xl_ext.sheet_names

                cs1, cs2 = st.columns(2)
                with cs1:
                    sheet_esc = st.selectbox("Sheet", sheets, key="m_sheet")
                with cs2:
                    hrow = int(st.number_input("Linha cabeçalho (0=primeira)", 0, 10, 0, key="m_hrow"))

                df_prev = xl_ext.parse(sheet_esc, header=hrow, nrows=3)
                df_prev.columns = [str(c).strip() for c in df_prev.columns]
                st.markdown("**Pré-visualização (3 linhas):**")
                st.dataframe(df_prev, use_container_width=True, hide_index=True)

                opcoes = ["— (não mapear)"] + list(df_prev.columns)
                campos_map = {
                    "designacao":     "Nome da entidade *(obrigatório)*",
                    "email":          "Email *(obrigatório)*",
                    "nome_dirigente": "Nome do dirigente",
                    "orgao_direcao":  "Cargo / Órgão",
                    "ministerio":     "Ministério",
                    "sigla_entidade": "Sigla",
                    "tipo_entidade":  "Tipo de entidade",
                    "contacto":       "Telefone",
                }
                mapa = {}
                cm1, cm2 = st.columns(2)
                for k, (campo, desc) in enumerate(campos_map.items()):
                    with (cm1 if k % 2 == 0 else cm2):
                        esc = st.selectbox(desc, opcoes, key=f"mc_{campo}")
                        if esc != "— (não mapear)":
                            mapa[campo] = esc

                cat_nome = st.text_input(
                    "Categoria a atribuir",
                    value=ext_name.replace(".xlsx","").replace(".XLSX","")[:40],
                    key="m_cat",
                )

                if st.button("🔀 Fundir com a base principal", type="primary",
                             use_container_width=True, key="btn_manual"):
                    if "designacao" not in mapa or "email" not in mapa:
                        st.error("Tens de mapear pelo menos 'Nome da entidade' e 'Email'.")
                    else:
                        try:
                            df_m = xl_ext.parse(sheet_esc, header=hrow)
                            df_m.columns = [str(c).strip() for c in df_m.columns]
                            novo = pd.DataFrame()
                            for campo in ["designacao","sigla_entidade","ministerio",
                                          "tipo_entidade","orgao_direcao","nome_dirigente",
                                          "email","contacto"]:
                                col_e = mapa.get(campo, "")
                                novo[campo] = (
                                    df_m[col_e].apply(limpar)
                                    if col_e and col_e in df_m.columns
                                    else ""
                                )
                            novo["email"]     = novo["email"].apply(email_ok)
                            novo["categoria"] = cat_nome
                            novo["fonte"]     = f"importado:{ext_name}"
                            novo = novo[novo["designacao"].str.len() > 2].copy()

                            if novo.empty:
                                st.warning("Nenhum registo válido encontrado (designação vazia ou em falta).")
                            else:
                                n_add, n_em = fundir_com_base(novo)
                                if n_add == 0:
                                    st.warning("Todos os registos já existem na base (emails duplicados).")
                                else:
                                    st.session_state["import_msg"] = {
                                        "texto": (
                                            f"✅ Importação manual concluída! "
                                            f"**{n_add:,}** registos adicionados à categoria **{cat_nome}**, "
                                            f"**{n_em:,}** com email. "
                                            f"Base agora tem **{len(st.session_state['df']):,}** registos."
                                        )
                                    }
                                    st.rerun()
                        except Exception as e:
                            st.error(f"Erro ao importar: {e}")

            except Exception as e:
                st.error(f"Erro a ler o ficheiro externo: {e}")


# ══════════════════════════════════════════════════════
# TAB 4 — DESCARREGAR
# ══════════════════════════════════════════════════════
with tab4:
    st.markdown("### 📥 Descarregar base de dados")

    df_dl = st.session_state["df"]
    n_tot  = len(df_dl)
    n_mail = (df_dl["email"].str.len() > 3).sum()
    n_unem = n_tot - n_mail
    st.info(f"Base actual: **{n_tot:,}** registos  |  **{n_mail:,}** com email  |  **{n_unem:,}** sem email")

    cd1, cd2, cd3 = st.columns(3)
    with cd1:
        st.markdown("**Excel completo** (uma sheet por categoria)")
        st.download_button(
            "⬇️ Excel completo",
            data=df_to_excel_bytes(df_dl),
            file_name=f"AP_Contactos_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
    with cd2:
        tem_filtro = bool(pesquisa or cats_sel or min_sel or so_email)
        lbl_xl_f = f"Excel filtrado ({len(df_f):,} registos)" if tem_filtro else "Excel filtrado (sem filtros activos)"
        st.markdown(f"**{lbl_xl_f}**")
        if not tem_filtro:
            st.warning("Activa pelo menos um filtro na barra lateral para usar esta opção.", icon="⚠️")
        else:
            st.download_button(
                f"⬇️ Excel filtrado ({len(df_f):,})",
                data=df_to_excel_bytes(df_f.drop(columns=["id"], errors="ignore")),
                file_name=f"AP_Contactos_filtrado_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )
    with cd3:
        st.markdown(f"**CSV filtrado** ({len(df_f):,} registos visíveis)")
        if not tem_filtro:
            st.warning("Activa pelo menos um filtro na barra lateral.", icon="⚠️")
        else:
            csv_b = (df_f.drop(columns=["id"], errors="ignore")
                     .to_csv(index=False, encoding="utf-8-sig")
                     .encode("utf-8-sig"))
            st.download_button(
                f"⬇️ CSV filtrado ({len(df_f):,})",
                data=csv_b,
                file_name=f"AP_Contactos_filtrado_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
            )

    st.markdown("---")
    st.markdown("**Cobertura de email por categoria:**")
    res = (df_dl.groupby("categoria")
           .agg(Total=("id","count"),
                Com_Email=("email", lambda x: (x.str.len() > 3).sum()))
           .reset_index())
    res.columns  = ["Categoria","Total","Com Email"]
    res["Sem Email"] = res["Total"] - res["Com Email"]
    res["Cobertura"] = (res["Com Email"] / res["Total"] * 100).round(1).astype(str) + "%"
    st.dataframe(
        res.sort_values("Total", ascending=False),
        use_container_width=True,
        hide_index=True,
    )
