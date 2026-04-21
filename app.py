"""
AP Contactos — Base de Contactos da Administração Pública Portuguesa
Estilo INA. Modelo entidades + dirigentes (com histórico).
Persistência: GitHub. SIOE manual. DRE como sugestões para aprovação.
"""
from __future__ import annotations

import io
import re
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st

from db import (
    COLS_ENTIDADE, COLS_DIRIGENTE,
    load_entidades, load_dirigentes, save_entidades, save_dirigentes,
    load_categorias, cat_nome, cat_cor, cat_ordem_map,
    gh_cfg, norm,
)
from ingest import importar_excel_ina, importar_sioe, merge_entidades, merge_dirigentes
from dre import sugestoes_a_partir_de_dre, carregar_pendentes, guardar_pendentes

st.set_page_config(page_title="AP Contactos · INA", page_icon="🟪", layout="wide")

# ═══════════════════════════════════════════════════
# AUTENTICAÇÃO
# ═══════════════════════════════════════════════════
_PWD = "INA#Contactos2026!"

def check_password():
    if st.session_state.get("autenticado"): return
    st.markdown("""
    <div style='max-width:480px;margin:80px auto 0;text-align:center;font-family:Inter,system-ui,sans-serif'>
      <div style='display:inline-grid;grid-template-columns:repeat(2,16px);gap:4px;margin-bottom:16px'>
        <div style='width:16px;height:16px;background:#701471'></div>
        <div style='width:16px;height:16px;background:#701471'></div>
        <div style='width:16px;height:16px;background:#701471'></div>
        <div style='width:16px;height:16px;background:transparent;border:1px solid #701471'></div>
      </div>
      <h1 style='color:#212121;font-weight:900;letter-spacing:-1px;margin:0;font-size:2.4rem'>ina</h1>
      <p style='color:#595959;margin:4px 0 28px;font-size:.85rem;letter-spacing:.5px;text-transform:uppercase'>
        Contactos da Administração Pública
      </p>
    </div>
    """, unsafe_allow_html=True)
    col = st.columns([1,1,1])[1]
    with col:
        pwd = st.text_input("", type="password", placeholder="Password",
                            label_visibility="collapsed")
        if st.button("Entrar", use_container_width=True, type="primary"):
            if pwd == _PWD:
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("Password incorrecta")
    st.stop()

check_password()

# ═══════════════════════════════════════════════════
# CSS — Identidade INA
# ═══════════════════════════════════════════════════
INA_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;900&display=swap');
:root {
  --ina-purple: #701471;
  --ina-magenta: #B1005D;
  --ina-orange: #FF8F2B;
  --ina-text: #212121;
  --ina-muted: #595959;
  --ina-bg: #FFFFFF;
  --ina-bg-alt: #F5F6F7;
  --ina-border: #E1E1E1;
}
html, body, [class*="css"], .stApp {
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
}
.stApp { background: var(--ina-bg) !important; color: var(--ina-text); }
section[data-testid="stSidebar"] { background: var(--ina-bg-alt) !important; border-right: 1px solid var(--ina-border); }
h1, h2, h3, h4 { color: var(--ina-text) !important; font-weight: 900 !important; letter-spacing: -.5px; }
h1 { font-size: 2.2rem !important; }
h2 { font-size: 1.5rem !important; }
.ina-header {
  display:flex; align-items:flex-start; gap:18px;
  padding:8px 0 20px; border-bottom: 1px solid var(--ina-border); margin-bottom: 24px;
}
.ina-logo {
  display:grid; grid-template-columns: repeat(2, 14px); gap:3px; flex-shrink:0; margin-top:6px;
}
.ina-logo span { width:14px; height:14px; background: var(--ina-purple); display:block; }
.ina-logo span.empty { background: transparent; border: 1px solid var(--ina-purple); }
.ina-eyebrow {
  color: var(--ina-purple); font-size:.7rem; font-weight:700;
  letter-spacing: 1.5px; text-transform: uppercase; display:flex; align-items:center; gap:8px;
}
.ina-eyebrow::before {
  content:''; display:inline-block; width:24px; height:1px; background: var(--ina-muted);
}
.ina-eyebrow::after {
  content:''; display:inline-block; width:8px; height:8px; background: var(--ina-purple);
}
.ina-title { font-weight:900; font-size: 1.8rem; line-height: 1.05; color: var(--ina-text); margin-top: 4px; }
.ina-sub   { color: var(--ina-muted); font-size:.95rem; margin-top: 2px; }
.kpi {
  background: white; border: 1px solid var(--ina-border); padding: 14px 18px; border-radius: 0;
  border-top: 3px solid var(--ina-purple);
}
.kpi .n { font-size: 1.9rem; font-weight: 900; color: var(--ina-text); line-height: 1; }
.kpi .l { font-size: .72rem; color: var(--ina-muted); margin-top: 6px; text-transform: uppercase; letter-spacing: 1px; }
.cat-pill {
  display:inline-block; padding:2px 8px; font-size:.7rem; font-weight:600;
  background: var(--ina-bg-alt); color: var(--ina-text); border-left: 3px solid var(--ina-purple);
}
.stTabs [data-baseweb="tab-list"] { gap: 0; border-bottom: 2px solid var(--ina-border); }
.stTabs [data-baseweb="tab"] {
  background: transparent !important; color: var(--ina-muted); font-weight: 700;
  padding: 10px 18px; border-radius: 0 !important;
}
.stTabs [aria-selected="true"] {
  color: var(--ina-purple) !important; border-bottom: 3px solid var(--ina-purple) !important; background: transparent !important;
}
.stButton > button[kind="primary"] {
  background: var(--ina-purple) !important; border: none !important;
  color: white !important; font-weight: 700 !important; border-radius: 0 !important;
}
.stButton > button[kind="primary"]:hover { background: #5e1060 !important; }
.stButton > button {
  border-radius: 0 !important; font-weight: 600 !important;
}
.stDownloadButton > button {
  border-radius: 0 !important; font-weight: 600 !important;
}
.stTextInput input, .stSelectbox > div, .stMultiSelect > div {
  border-radius: 0 !important;
}
.email-mono {
  background: var(--ina-bg-alt); border: 1px solid var(--ina-border); padding: 12px 16px;
  font-family: 'JetBrains Mono', Consolas, monospace; font-size:.82rem;
  max-height: 320px; overflow-y: auto; white-space: pre-wrap; word-break: break-all;
}
hr { border-color: var(--ina-border) !important; }
.dirigente-historico { color: #999 !important; text-decoration: line-through; }
.alerta-banner {
  background: #FFF3E0; border-left: 4px solid var(--ina-orange);
  padding: 10px 14px; margin: 8px 0; font-size: .9rem;
}
.ok-banner {
  background: #E8F5E9; border-left: 4px solid #2E7D32;
  padding: 10px 14px; margin: 8px 0; font-size: .9rem;
}
</style>
"""
st.markdown(INA_CSS, unsafe_allow_html=True)

def render_header(titulo: str, sub: str = ""):
    st.markdown(f"""
    <div class='ina-header'>
      <div class='ina-logo'>
        <span></span><span></span><span></span><span class='empty'></span>
      </div>
      <div>
        <div class='ina-eyebrow'>Instituto Nacional de Administração</div>
        <div class='ina-title'>{titulo}</div>
        <div class='ina-sub'>{sub}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

def kpi(num, label):
    st.markdown(f"<div class='kpi'><div class='n'>{num}</div><div class='l'>{label}</div></div>",
                unsafe_allow_html=True)

# ═══════════════════════════════════════════════════
# ESTADO + LOAD INICIAL
# ═══════════════════════════════════════════════════
defaults = {
    "ent": None, "dir": None,
    "sha_ent": None, "sha_dir": None,
    "carregado": False,
    "alterado_ent": False, "alterado_dir": False,
    "msg": None, "log_import": None,
    "_file_key": None, "_ext_key": None, "_ext_bytes": None, "_ext_name": None,
}
for k, v in defaults.items():
    if k not in st.session_state: st.session_state[k] = v

if not st.session_state["carregado"] and gh_cfg():
    with st.spinner("A carregar base permanente do GitHub..."):
        e, sha_e = load_entidades()
        d, sha_d = load_dirigentes()
        st.session_state["ent"] = e
        st.session_state["dir"] = d
        st.session_state["sha_ent"] = sha_e
        st.session_state["sha_dir"] = sha_d
        st.session_state["carregado"] = True

ent = st.session_state["ent"] if st.session_state["ent"] is not None else pd.DataFrame(columns=COLS_ENTIDADE)
dir_ = st.session_state["dir"] if st.session_state["dir"] is not None else pd.DataFrame(columns=COLS_DIRIGENTE)

# ═══════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div style='display:flex;align-items:center;gap:10px;padding:8px 0 14px;border-bottom:1px solid #e1e1e1'>
      <div style='display:grid;grid-template-columns:repeat(2,10px);gap:2px'>
        <div style='width:10px;height:10px;background:#701471'></div>
        <div style='width:10px;height:10px;background:#701471'></div>
        <div style='width:10px;height:10px;background:#701471'></div>
        <div style='width:10px;height:10px;background:transparent;border:1px solid #701471'></div>
      </div>
      <div>
        <div style='font-weight:900;font-size:1.1rem;color:#212121'>ina · contactos</div>
        <div style='color:#595959;font-size:.7rem;letter-spacing:.5px;text-transform:uppercase'>AP Portuguesa</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    if gh_cfg():
        st.markdown(f"<div class='ok-banner'>🔗 Persistência GitHub activa<br><span style='color:#666;font-size:.78rem'>{len(ent):,} entidades · {len(dir_):,} dirigentes</span></div>",
                    unsafe_allow_html=True)
    else:
        st.markdown("<div class='alerta-banner'>⚠ GitHub não configurado.<br>Define <code>GITHUB_TOKEN</code> e <code>GITHUB_REPO</code> nos Secrets.</div>",
                    unsafe_allow_html=True)

    # Botão único de gravação
    pendente = st.session_state["alterado_ent"] or st.session_state["alterado_dir"]
    if pendente and gh_cfg():
        if st.button("💾 Guardar alterações", type="primary", use_container_width=True):
            ok_e = ok_d = True; msg = ""
            if st.session_state["alterado_ent"]:
                ok_e, m, sha_e = save_entidades(st.session_state["ent"], st.session_state["sha_ent"])
                if ok_e:
                    st.session_state["sha_ent"] = sha_e
                    st.session_state["alterado_ent"] = False
                    msg += f"Entidades: {m}. "
                else:
                    msg += f"❌ Entidades: {m}. "
            if st.session_state["alterado_dir"]:
                ok_d, m, sha_d = save_dirigentes(st.session_state["dir"], st.session_state["sha_dir"])
                if ok_d:
                    st.session_state["sha_dir"] = sha_d
                    st.session_state["alterado_dir"] = False
                    msg += f"Dirigentes: {m}. "
                else:
                    msg += f"❌ Dirigentes: {m}. "
            if ok_e and ok_d:
                st.session_state["msg"] = ("ok", "✅ Tudo guardado.")
            else:
                st.session_state["msg"] = ("err", msg)
            st.rerun()

    st.markdown("---")
    st.markdown("##### 📂 Importar dados")
    up = st.file_uploader("Excel INA / SIOE (.xlsx)", type=["xlsx"], label_visibility="collapsed")
    if up is not None:
        fkey = f"{up.name}_{up.size}"
        if st.session_state["_file_key"] != fkey:
            st.session_state["_file_key"] = fkey
            st.session_state["_ext_bytes"] = up.getvalue()
            st.session_state["_ext_name"]  = up.name

    st.markdown("---")
    st.markdown("##### 🔍 Filtros")
    pesquisa = st.text_input("Pesquisar", "", placeholder="entidade, sigla, nome, email…")
    mostrar_historico = st.checkbox("Mostrar dirigentes históricos", value=False)
    cats = load_categorias()["categorias"]
    cats_opt = [(c["id"], c["nome"]) for c in sorted(cats, key=lambda x: x["ordem"])]
    cat_sel_ids = st.multiselect("Categoria",
        options=[c[0] for c in cats_opt],
        format_func=lambda i: dict(cats_opt).get(i, i))
    so_email = st.checkbox("Só com email", value=False)

    st.markdown("---")
    if st.button("↻ Recarregar do GitHub", use_container_width=True):
        st.session_state["carregado"] = False
        st.rerun()

# ═══════════════════════════════════════════════════
# CABEÇALHO + KPIs
# ═══════════════════════════════════════════════════
render_header("Contactos da Administração Pública",
              "Base operacional · 2 ou 3 utilizadores · persistência GitHub")

if st.session_state["msg"]:
    tipo, txt = st.session_state["msg"]
    klass = "ok-banner" if tipo == "ok" else "alerta-banner"
    st.markdown(f"<div class='{klass}'>{txt}</div>", unsafe_allow_html=True)
    if st.button("✕ Fechar mensagem", key="close_msg"):
        st.session_state["msg"] = None; st.rerun()

# Filtros aplicados
def aplicar_filtros(ent_df, dir_df):
    e = ent_df.copy()
    d = dir_df.copy()
    if not mostrar_historico and len(d):
        d = d[d["fim"].fillna("") == ""]
    if cat_sel_ids:
        e = e[e["categoria_id"].isin(cat_sel_ids)]
        d = d[d["entity_id"].isin(e["entity_id"])]
    if so_email and len(d):
        d = d[d["email"].str.len() > 3]
        e = e[e["entity_id"].isin(d["entity_id"])]
    if pesquisa:
        p = norm(pesquisa)
        m_e = (e["designacao"].apply(norm).str.contains(p, na=False) |
               e["siglas"].apply(norm).str.contains(p, na=False) |
               e["ministerio"].apply(norm).str.contains(p, na=False))
        m_d = (d["nome"].apply(norm).str.contains(p, na=False) |
               d["email"].apply(norm).str.contains(p, na=False) |
               d["cargo"].apply(norm).str.contains(p, na=False))
        ent_ids = set(e[m_e]["entity_id"]) | set(d[m_d]["entity_id"])
        e = ent_df[ent_df["entity_id"].isin(ent_ids)]
        d = dir_df[dir_df["entity_id"].isin(ent_ids)]
        if not mostrar_historico:
            d = d[d["fim"].fillna("") == ""]
    return e, d

ent_f, dir_f = aplicar_filtros(ent, dir_)
dir_ativos = dir_f[dir_f["fim"].fillna("") == ""] if len(dir_f) else dir_f
emails_validos = dir_ativos[dir_ativos["email"].str.len() > 3] if len(dir_ativos) else dir_ativos

c1, c2, c3, c4 = st.columns(4)
with c1: kpi(f"{len(ent_f):,}", "Entidades")
with c2: kpi(f"{len(dir_ativos):,}", "Dirigentes activos")
with c3: kpi(f"{len(emails_validos):,}", "Com email")
with c4: kpi(f"{ent_f['ministerio'].replace('',pd.NA).dropna().nunique():,}", "Ministérios")

# ═══════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📋 Tabela", "👤 Por entidade",
    "✉️ Copiar emails", "📥 Importar / Exportar",
    "🔔 Revisão (DRE)", "📚 Histórico",
])

# ─────────── TAB 1: TABELA ───────────
with tab1:
    if len(ent_f) == 0:
        st.info("Nenhuma entidade na base ainda. Importa o teu Excel INA na barra lateral → Importar / Exportar.")
    else:
        # Juntar entidade + dirigente activo principal
        d_act = dir_f[dir_f["fim"].fillna("") == ""] if len(dir_f) else pd.DataFrame(columns=COLS_DIRIGENTE)
        # primeiro dirigente por entidade (ordem por inserção)
        d_first = d_act.drop_duplicates(subset=["entity_id"], keep="first") if len(d_act) else d_act
        merged = ent_f.merge(d_first[["entity_id","nome","cargo","email","telefone"]],
                             on="entity_id", how="left")
        merged["categoria"] = merged["categoria_id"].apply(cat_nome)
        cols_v = {
            "categoria":"Categoria", "siglas":"Sigla", "designacao":"Entidade",
            "ministerio":"Ministério", "cargo":"Cargo",
            "nome":"Dirigente", "email":"Email", "telefone":"Telefone",
        }
        cols_ok = [c for c in cols_v if c in merged.columns]
        st.dataframe(merged[cols_ok].rename(columns=cols_v),
                     use_container_width=True, height=560, hide_index=True)
        st.caption(f"{len(merged):,} entidades · 1 dirigente activo por entidade · "
                   f"{(merged['email'].fillna('').str.len()>3).sum():,} com email")

# ─────────── TAB 2: POR ENTIDADE ───────────
with tab2:
    if len(ent_f) == 0:
        st.info("Sem entidades.")
    else:
        designacoes = ent_f.sort_values("designacao")["designacao"].tolist()
        ent_escolhida = st.selectbox("Escolhe uma entidade", designacoes)
        if ent_escolhida:
            row = ent_f[ent_f["designacao"] == ent_escolhida].iloc[0]
            eid = row["entity_id"]
            colA, colB = st.columns([2, 1])
            with colA:
                st.markdown(f"### {row['designacao']}")
                if row.get("siglas"):
                    st.markdown(f"<span class='cat-pill' style='border-left-color:{cat_cor(row['categoria_id'])}'>{cat_nome(row['categoria_id'])}</span> · <strong>{row['siglas']}</strong>", unsafe_allow_html=True)
                else:
                    st.markdown(f"<span class='cat-pill' style='border-left-color:{cat_cor(row['categoria_id'])}'>{cat_nome(row['categoria_id'])}</span>", unsafe_allow_html=True)
                if row.get("ministerio"): st.write(f"**Ministério:** {row['ministerio']}")
                if row.get("sioe_code"):  st.write(f"**Código SIOE:** {row['sioe_code']}")
                if row.get("website"):    st.write(f"**Website:** {row['website']}")
            with colB:
                st.metric("Dirigentes activos",
                          len(dir_[(dir_["entity_id"]==eid) & (dir_["fim"].fillna("")=="")]))
                st.metric("Histórico",
                          len(dir_[(dir_["entity_id"]==eid) & (dir_["fim"].fillna("")!="")]))

            st.markdown("#### Dirigentes activos")
            d_e = dir_[dir_["entity_id"] == eid]
            d_act = d_e[d_e["fim"].fillna("") == ""]
            if len(d_act):
                st.dataframe(d_act[["cargo","nome","email","telefone","fonte"]]
                             .rename(columns={"cargo":"Cargo","nome":"Nome","email":"Email","telefone":"Telefone","fonte":"Fonte"}),
                             use_container_width=True, hide_index=True)
            else:
                st.caption("Sem dirigentes activos registados.")

            d_hist = d_e[d_e["fim"].fillna("") != ""]
            if len(d_hist):
                with st.expander(f"📚 Histórico ({len(d_hist)})"):
                    st.dataframe(d_hist[["cargo","nome","email","fim","fonte"]]
                                 .rename(columns={"cargo":"Cargo","nome":"Nome","email":"Email","fim":"Saída","fonte":"Fonte"}),
                                 use_container_width=True, hide_index=True)

# ─────────── TAB 3: COPIAR EMAILS ───────────
with tab3:
    st.markdown("### Exportar emails do filtro actual")
    base_emails = emails_validos.drop_duplicates(subset=["email"]) if len(emails_validos) else emails_validos
    n = len(base_emails)
    if n == 0:
        st.info("Sem emails no filtro actual.")
    else:
        st.write(f"**{n:,}** emails únicos")
        c1, c2 = st.columns([2,1])
        with c1:
            fmt = st.radio("Formato", ["Um por linha","Separados por ; (BCC Outlook)","CSV (Nome,Email,Cargo,Entidade)"], horizontal=False)
        with c2:
            inc_nome = st.checkbox("Incluir nome", value=True)
        emails = sorted(base_emails["email"].tolist())
        if fmt.startswith("Um"):
            txt = "\n".join(f"{r['nome']} <{r['email']}>" if inc_nome and str(r.get('nome','')).strip() else r['email']
                            for _, r in base_emails.sort_values("email").iterrows())
        elif fmt.startswith("Separados"):
            if inc_nome:
                txt = "; ".join(f"{r['nome']} <{r['email']}>" if str(r.get('nome','')).strip() else r['email']
                                for _, r in base_emails.sort_values("email").iterrows())
            else:
                txt = "; ".join(emails)
        else:
            joined = base_emails.merge(ent[["entity_id","designacao"]], on="entity_id", how="left")
            txt = "Nome,Email,Cargo,Entidade\n" + "\n".join(
                '"' + str(r.get("nome","")) + '",' + r["email"] + ',"' + str(r.get("cargo","")) + '","' + str(r.get("designacao","")) + '"'
                for _, r in joined.sort_values("email").iterrows())

        st.markdown(f"<div class='email-mono'>{txt[:8000].replace('<','&lt;').replace('>','&gt;')}</div>", unsafe_allow_html=True)
        if len(txt) > 8000: st.caption(f"…(pré-visualização truncada nos primeiros 8000 caracteres; o ficheiro completo tem {len(txt):,} caracteres)")

        st.download_button("⬇ Descarregar (.txt)" if not fmt.startswith("CSV") else "⬇ Descarregar (.csv)",
            data=txt.encode("utf-8-sig"),
            file_name=f"AP_Contactos_emails_{datetime.now().strftime('%Y%m%d_%H%M')}." +
                      ("csv" if fmt.startswith("CSV") else "txt"),
            mime="text/csv" if fmt.startswith("CSV") else "text/plain",
            type="primary", use_container_width=True)

# ─────────── TAB 4: IMPORTAR / EXPORTAR ───────────
with tab4:
    st.markdown("### Importar")
    if st.session_state["_ext_bytes"] is None:
        st.info("Carrega um ficheiro Excel na barra lateral.\n\n"
                "**Tipos suportados:**\n"
                "- 🟪 **Excel INA** (multi-sheet por categoria) — ex: Base_de_dados_17_04_ATUALIZADO.xlsx\n"
                "- 🟧 **SIOE** (ExportResultadosPesquisa*.xlsx) — exportação manual do sioe.pt")
    else:
        nm = st.session_state["_ext_name"]
        st.write("**Ficheiro:** " + nm)
        eh_sioe = "export" in nm.lower() and "pesquisa" in nm.lower()
        try:
            xl = pd.ExcelFile(io.BytesIO(st.session_state["_ext_bytes"]))
            sheets = xl.sheet_names
            st.caption(f"{len(sheets)} sheets detectadas: {', '.join(sheets[:8])}{'…' if len(sheets)>8 else ''}")
        except Exception as e:
            st.error(f"Erro ao ler Excel: {e}")
            sheets = []

        cI, cII = st.columns(2)
        with cI:
            if st.button("🟪 Importar como Excel INA", type="primary", use_container_width=True, disabled=eh_sioe):
                with st.spinner("A processar Excel INA…"):
                    try:
                        ne, nd, log = importar_excel_ina(st.session_state["_ext_bytes"])
                        ent_m, rel_e = merge_entidades(ent, ne)
                        dir_m, rel_d = merge_dirigentes(dir_, nd, ent, ne)
                        st.session_state["ent"] = ent_m
                        st.session_state["dir"] = dir_m
                        st.session_state["alterado_ent"] = True
                        st.session_state["alterado_dir"] = True
                        st.session_state["log_import"] = log
                        st.session_state["msg"] = ("ok",
                            f"✅ Excel INA processado · {log['entidades_distintas']:,} entidades distintas, "
                            f"{log['dirigentes_distintos']:,} dirigentes distintos no ficheiro. "
                            f"<br>Merge: <strong>{rel_e['adicionadas']:,}</strong> entidades novas, "
                            f"<strong>{rel_e['atualizadas']:,}</strong> campos atualizados; "
                            f"<strong>{rel_d['adicionados']:,}</strong> dirigentes novos, "
                            f"<strong>{rel_d['sucessoes']:,}</strong> sucessões, "
                            f"<strong>{rel_d['duplicados']:,}</strong> duplicados ignorados. "
                            f"<br>💾 Carrega <strong>Guardar alterações</strong> na barra lateral.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro: {e}")
        with cII:
            if st.button("🟧 Importar como SIOE", type="primary", use_container_width=True, disabled=not eh_sioe):
                with st.spinner("A processar SIOE…"):
                    try:
                        ne, nd, log = importar_sioe(st.session_state["_ext_bytes"])
                        ent_m, rel_e = merge_entidades(ent, ne)
                        dir_m, rel_d = merge_dirigentes(dir_, nd, ent, ne)
                        st.session_state["ent"] = ent_m
                        st.session_state["dir"] = dir_m
                        st.session_state["alterado_ent"] = True
                        st.session_state["alterado_dir"] = True
                        st.session_state["msg"] = ("ok",
                            f"✅ SIOE processado · {log['entidades']:,} entidades, {log['dirigentes']:,} dirigentes. "
                            f"Merge: {rel_e['adicionadas']} novas, {rel_d['adicionados']} dirigentes novos, "
                            f"{rel_d['sucessoes']} sucessões. 💾 Carrega Guardar.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro: {e}")

        if st.session_state.get("log_import"):
            with st.expander("📋 Log da última importação"):
                log = st.session_state["log_import"]
                st.write("**Linhas processadas:** " + str(log.get('linhas_processadas','?')))
                st.write("**Sheets OK:**")
                for s in log.get("sheets_ok", []): st.caption("✓ " + s)
                if log.get("sheets_ign"):
                    st.write("**Ignoradas:**")
                    for s in log["sheets_ign"]: st.caption("× " + s)

        if st.button("🗑 Remover ficheiro", key="rm_ext"):
            st.session_state["_file_key"] = None
            st.session_state["_ext_bytes"] = None
            st.session_state["_ext_name"] = None
            st.rerun()

    st.markdown("---")
    st.markdown("### Exportar")
    if len(ent) == 0:
        st.caption("Nada para exportar.")
    else:
        merged_full = ent.merge(
            dir_[dir_["fim"].fillna("")==""].drop_duplicates(subset=["entity_id"], keep="first")[
                ["entity_id","nome","cargo","email","telefone"]],
            on="entity_id", how="left")
        merged_full["categoria"] = merged_full["categoria_id"].apply(cat_nome)
        cols_x = ["categoria","siglas","designacao","ministerio","cargo","nome","email","telefone","sioe_code","website"]
        cols_x = [c for c in cols_x if c in merged_full.columns]
       def _safe_sheet_name(name, used):
            s = re.sub(r'[\\/*?:\[\]]', ' ', str(name or '')).strip()
            s = re.sub(r'\s+', ' ', s)
            if not s:
                s = "Sem categoria"
            if s.lower() == "history":
                s = "Historico"
            s = s[:31]
            base, i = s, 2
            while s in used:
                suf = f" ({i})"
                s = (base[:31 - len(suf)]) + suf
                i += 1
            used.add(s)
            return s

        buf = io.BytesIO()
        cmap = cat_ordem_map()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            used_names = set()
            merged_full[cols_x].to_excel(
                w, sheet_name=_safe_sheet_name("Todos", used_names), index=False)
            for cid in sorted(merged_full["categoria_id"].dropna().unique(),
                              key=lambda c: cmap.get(c, 99)):
                sub = merged_full[merged_full["categoria_id"] == cid][cols_x]
                if len(sub):
                    sub.to_excel(
                        w,
                        sheet_name=_safe_sheet_name(cat_nome(cid), used_names),
                        index=False)
        st.download_button("⬇ Excel completo (sheet por categoria)",
            data=buf.getvalue(),
            file_name=f"AP_Contactos_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary")
        st.download_button("⬇ entidades.csv (raw)",
            data=ent.to_csv(index=False).encode("utf-8-sig"),
            file_name="entidades.csv", mime="text/csv")
        st.download_button("⬇ dirigentes.csv (raw)",
            data=dir_.to_csv(index=False).encode("utf-8-sig"),
            file_name="dirigentes.csv", mime="text/csv")

# ─────────── TAB 5: REVISÃO (DRE) ───────────
with tab5:
    st.markdown("### 🔔 Revisão de sugestões automáticas (DRE)")
    st.caption("As sugestões são candidatos detectados no Diário da República. "
               "Nada é inserido sem a tua aprovação.")

    col_a, col_b = st.columns([1,1])
    with col_a:
        dias = st.number_input("Procurar nas últimas (dias)", 7, 90, 30, 7)
    with col_b:
        if st.button("🔍 Procurar agora no DRE", type="primary", use_container_width=True):
            with st.spinner("A consultar dre.pt…"):
                novas = sugestoes_a_partir_de_dre(ent, dias=int(dias))
                if novas:
                    pendentes = carregar_pendentes()
                    ids_existentes = {p["id"] for p in pendentes}
                    pendentes.extend([n for n in novas if n["id"] not in ids_existentes])
                    if gh_cfg():
                        ok, m = guardar_pendentes(pendentes)
                        if ok:
                            st.success(f"{len(novas)} sugestões adicionadas à fila.")
                        else:
                            st.error(f"Falha a guardar pendentes: {m}")
                    else:
                        st.warning("GitHub não configurado — sugestões só nesta sessão.")
                        st.session_state["_pendentes_local"] = pendentes
                else:
                    st.info("Nenhuma sugestão nova encontrada (ou DRE indisponível).")

    pendentes = carregar_pendentes() if gh_cfg() else st.session_state.get("_pendentes_local", [])
    pendentes_abertos = [p for p in pendentes if p.get("estado") == "pendente"]
    st.markdown(f"**{len(pendentes_abertos)} sugestões pendentes**")

    if not pendentes_abertos:
        st.caption("Nada à espera de revisão.")
    else:
        for p in pendentes_abertos[:30]:
            with st.container():
                cA, cB, cC = st.columns([3, 1, 1])
                with cA:
                    ent_nome = ""
                    if p.get("entity_id"):
                        match = ent[ent["entity_id"] == p["entity_id"]]
                        if len(match): ent_nome = match.iloc[0]["designacao"]
                    st.markdown(f"**{p.get('tipo_acao','?').upper()}** · {ent_nome or p.get('entity_id','—')}")
                    st.caption(p.get("titulo","")[:200])
                    st.caption(f"Evidência: {p.get('evidencia','')} · [DRE]({p.get('url','#')})")
                with cB:
                    if st.button("✅ Aprovar", key="ap_" + p["id"]):
                        p["estado"] = "aprovado"
                        if gh_cfg(): guardar_pendentes(pendentes)
                        st.rerun()
                with cC:
                    if st.button("❌ Rejeitar", key="rj_" + p["id"]):
                        p["estado"] = "rejeitado"
                        if gh_cfg(): guardar_pendentes(pendentes)
                        st.rerun()
                st.markdown("---")

# ─────────── TAB 6: HISTÓRICO ───────────
with tab6:
    st.markdown("### 📚 Histórico de dirigentes")
    st.caption("Dirigentes que já saíram do cargo. Sucessões automáticas marcam-se aqui.")
    d_h = dir_[dir_["fim"].fillna("") != ""].copy()
    if len(d_h) == 0:
        st.info("Sem histórico ainda. Quando um dirigente é substituído por outra importação, fica aqui.")
    else:
        d_h = d_h.merge(ent[["entity_id","designacao","categoria_id"]], on="entity_id", how="left")
        d_h["Categoria"] = d_h["categoria_id"].apply(cat_nome)
        st.dataframe(
            d_h[["Categoria","designacao","cargo","nome","email","fim","fonte"]]
              .rename(columns={"designacao":"Entidade","cargo":"Cargo","nome":"Nome","email":"Email","fim":"Saída","fonte":"Fonte"})
              .sort_values("Saída", ascending=False),
            use_container_width=True, hide_index=True, height=520)
        st.download_button("⬇ Exportar histórico (.csv)",
            data=d_h.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"AP_historico_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv")
