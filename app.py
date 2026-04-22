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
# HELPER: nome de sheet seguro para Excel/openpyxl
# ═══════════════════════════════════════════════════
def safe_sheet_name(name, used):
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


# ═══════════════════════════════════════════════════
# AUTENTICAÇÃO
# ═══════════════════════════════════════════════════
_PWD = "INA#Contactos2026!"

def check_password():
    if st.session_state.get("autenticado"):
        return
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
    col = st.columns([1, 1, 1])[1]
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
.ina-eyebrow::before { content:''; display:inline-block; width:24px; height:1px; background: var(--ina-muted); }
.ina-eyebrow::after { content:''; display:inline-block; width:8px; height:8px; background: var(--ina-purple); }
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
.stButton > button { border-radius: 0 !important; font-weight: 600 !important; }
.stDownloadButton > button { border-radius: 0 !important; font-weight: 600 !important; }
.stTextInput input, .stSelectbox > div, .stMultiSelect > div { border-radius: 0 !important; }
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
    st.markdown(
        f"<div class='kpi'><div class='n'>{num}</div><div class='l'>{label}</div></div>",
        unsafe_allow_html=True,
    )


# ═══════════════════════════════════════════════════
# ESTADO + LOAD INICIAL
# ═══════════════════════════════════════════════════
defaults = {
    "ent": None, "dir": None,
    "sha_ent": None, "sha_dir": None,
    "carregado": False,
    "_load_retry": 0,
    "alterado_ent": False, "alterado_dir": False,
    "msg": None, "log_import": None,
    "_file_key": None, "_ext_key": None, "_ext_bytes": None, "_ext_name": None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

storage_label = "GitHub" if gh_cfg() else "base local"
storage_subtitle = "persistência GitHub" if gh_cfg() else "persistência local"

if not st.session_state["carregado"]:
    with st.spinner(f"A carregar {storage_label}..."):
        e, sha_e = load_entidades()
        d, sha_d = load_dirigentes()
        # Se o GitHub estiver ativo mas o arranque vier vazio, tentar mais uma vez
        # em vez de prender a sessão num falso "0 entidades".
        if gh_cfg() and len(e) == 0 and len(d) == 0 and st.session_state["_load_retry"] < 1:
            st.session_state["_load_retry"] += 1
            st.rerun()
        st.session_state["ent"] = e
        st.session_state["dir"] = d
        st.session_state["sha_ent"] = sha_e
        st.session_state["sha_dir"] = sha_d
        st.session_state["carregado"] = True
        st.session_state["_load_retry"] = 0

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
        st.markdown(
            f"<div class='ok-banner'>🔗 Persistência GitHub activa<br>"
            f"<span style='color:#666;font-size:.78rem'>{len(ent):,} entidades · {len(dir_):,} dirigentes</span></div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"<div class='ok-banner'>💾 Base local activa<br>"
            f"<span style='color:#666;font-size:.78rem'>{len(ent):,} entidades · {len(dir_):,} dirigentes</span></div>",
            unsafe_allow_html=True,
        )

    pendente = st.session_state["alterado_ent"] or st.session_state["alterado_dir"]
    if pendente:
        if st.button("💾 Guardar alterações", type="primary", use_container_width=True):
            ok_e = ok_d = True
            msg = ""
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
            st.session_state["_ext_name"] = up.name

    st.markdown("---")
    st.markdown("##### 🔍 Filtros")
    pesquisa = st.text_input("Pesquisar", "", placeholder="entidade, sigla, nome, email…")
    mostrar_historico = st.checkbox("Mostrar dirigentes históricos", value=False)
    cats = load_categorias()["categorias"]
    cats_opt = [(c["id"], c["nome"]) for c in sorted(cats, key=lambda x: x["ordem"])]
    cat_sel_ids = st.multiselect(
        "Categoria",
        options=[c[0] for c in cats_opt],
        format_func=lambda i: dict(cats_opt).get(i, i),
    )
    so_email = st.checkbox("Só com email", value=False)

    st.markdown("---")
    if st.button("↻ Recarregar dados", use_container_width=True):
        st.session_state["carregado"] = False
        st.rerun()


# ═══════════════════════════════════════════════════
# CABEÇALHO + KPIs
# ═══════════════════════════════════════════════════
render_header(
    "Contactos da Administração Pública",
    f"Base operacional · 2 ou 3 utilizadores · {storage_subtitle}",
)

if st.session_state["msg"]:
    tipo, txt = st.session_state["msg"]
    klass = "ok-banner" if tipo == "ok" else "alerta-banner"
    st.markdown(f"<div class='{klass}'>{txt}</div>", unsafe_allow_html=True)
    if st.button("✕ Fechar mensagem", key="close_msg"):
        st.session_state["msg"] = None
        st.rerun()


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
        m_e = (
            e["designacao"].apply(norm).str.contains(p, na=False)
            | e["siglas"].apply(norm).str.contains(p, na=False)
            | e["ministerio"].apply(norm).str.contains(p, na=False)
        )
        m_d = (
            d["nome"].apply(norm).str.contains(p, na=False)
            | d["email"].apply(norm).str.contains(p, na=False)
            | d["cargo"].apply(norm).str.contains(p, na=False)
        )
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
with c1:
    kpi(f"{len(ent_f):,}", "Entidades")
with c2:
    kpi(f"{len(dir_ativos):,}", "Dirigentes activos")
with c3:
    kpi(f"{len(emails_validos):,}", "Com email")
with c4:
    kpi(f"{ent_f['ministerio'].replace('', pd.NA).dropna().nunique():,}", "Ministérios")


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
        d_act = dir_f[dir_f["fim"].fillna("") == ""] if len(dir_f) else pd.DataFrame(columns=COLS_DIRIGENTE)
        d_first = d_act.drop_duplicates(subset=["entity_id"], keep="first") if len(d_act) else d_act
        merged = ent_f.merge(
            d_first[["entity_id", "nome", "cargo", "email", "telefone"]],
            on="entity_id", how="left",
        )
        merged["categoria"] = merged["categoria_id"].apply(cat_nome)
        cols_v = {
            "categoria": "Categoria", "siglas": "Sigla", "designacao": "Entidade",
            "ministerio": "Ministério", "cargo": "Cargo",
            "nome": "Dirigente", "email": "Email", "telefone": "Telefone",
        }
        cols_ok = [c for c in cols_v if c in merged.columns]
        st.dataframe(
            merged[cols_ok].rename(columns=cols_v),
            use_container_width=True, height=560, hide_index=True,
        )
        st.caption(
            f"{len(merged):,} entidades · 1 dirigente activo por entidade · "
            f"{(merged['email'].fillna('').str.len() > 3).sum():,} com email"
        )

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
                    st.markdown(
                        f"<span class='cat-pill' style='border-left-color:{cat_cor(row['categoria_id'])}'>"
                        f"{cat_nome(row['categoria_id'])}</span> · <strong>{row['siglas']}</strong>",
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        f"<span class='cat-pill' style='border-left-color:{cat_cor(row['categoria_id'])}'>"
                        f"{cat_nome(row['categoria_id'])}</span>",
                        unsafe_allow_html=True,
                    )
                if row.get("ministerio"):
                    st.write(f"**Ministério:** {row['ministerio']}")
                if row.get("sioe_code"):
                    st.write(f"**Código SIOE:** {row['sioe_code']}")
                if row.get("website"):
                    st.write(f"**Website:** {row['website']}")
            with colB:
                st.metric(
                    "Dirigentes activos",
                    len(dir_[(dir_["entity_id"] == eid) & (dir_["fim"].fillna("") == "")]),
                )
                st.metric(
                    "Histórico",
                    len(dir_[(dir_["entity_id"] == eid) & (dir_["fim"].fillna("") != "")]),
                )

            st.markdown("#### Dirigentes activos")
            d_e = dir_[dir_["entity_id"] == eid]
            d_act = d_e[d_e["fim"].fillna("") == ""]
            if len(d_act):
                st.dataframe(
                    d_act[["cargo", "nome", "email", "telefone", "fonte"]].rename(
                        columns={"cargo": "Cargo", "nome": "Nome", "email": "Email",
                                 "telefone": "Telefone", "fonte": "Fonte"}
                    ),
                    use_container_width=True, hide_index=True,
                )
            else:
                st.caption("Sem dirigentes activos registados.")

            d_hist = d_e[d_e["fim"].fillna("") != ""]
            if len(d_hist):
                with st.expander(f"📚 Histórico ({len(d_hist)})"):
                    st.dataframe(
                        d_hist[["cargo", "nome", "email", "fim", "fonte"]].rename(
                            columns={"cargo": "Cargo", "nome": "Nome", "email": "Email",
                                     "fim": "Saída", "fonte": "Fonte"}
                        ),
                        use_container_width=True, hide_index=True,
                    )

# ─────────── TAB 3: COPIAR EMAILS ───────────
with tab3:
    st.markdown("### Exportar emails do filtro actual")
    base_emails = emails_validos.drop_duplicates(subset=["email"]) if len(emails_validos) else emails_validos
    n = len(base_emails)
    if n == 0:
        st.info("Sem emails no filtro actual.")
    else:
        st.write(f"**{n:,}** emails únicos")
        c1, c2 = st.columns([2, 1])
        with c1:
            fmt = st.radio(
                "Formato",
                ["Um por linha", "Separados por ; (BCC Outlook)", "CSV (Nome,Email,Cargo,Entidade)"],
                horizontal=False,
            )
        with c2:
            inc_nome = st.checkbox("Incluir nome", value=True)

        emails = sorted(base_emails["email"].tolist())
        if fmt.startswith("Um"):
            txt = "\n".join(
                f"{r['nome']} <{r['email']}>" if inc_nome and str(r.get('nome', '')).strip() else r['email']
                for _, r in base_emails.sort_values("email").iterrows()
            )
        elif fmt.startswith("Separados"):
            if inc_nome:
                txt = "; ".join(
                    f"{r['nome']} <{r['email']}>" if str(r.get('nome', '')).strip() else r['email']
                    for _, r in base_emails.sort_values("email").iterrows()
                )
            else:
                txt = "; ".join(emails)
        else:
            joined = base_emails.merge(ent[["entity_id", "designacao"]], on="entity_id", how="left")
            txt = "Nome,Email,Cargo,Entidade\n" + "\n".join(
                '"' + str(r.get("nome", "")) + '",' + r["email"] + ',"' + str(r.get("cargo", "")) +
                '","' + str(r.get("designacao", "")) + '"'
                for _, r in joined.sort_values("email").iterrows()
            )

        st.markdown(
            f"<div class='email-mono'>{txt[:8000].replace('<', '&lt;').replace('>', '&gt;')}</div>",
            unsafe_allow_html=True,
        )
        if len(txt) > 8000:
            st.caption(
                f"…(pré-visualização truncada nos primeiros 8000 caracteres; "
                f"o ficheiro completo tem {len(txt):,} caracteres)"
            )
        st.download_button(
            "⬇ Descarregar (.txt)" if not fmt.startswith("CSV") else "⬇ Descarregar (.csv)",
            data=txt.encode("utf-8-sig"),
            file_name=f"AP_Contactos_emails_{datetime.now().strftime('%Y%m%d_%H%M')}."
                      + ("csv" if fmt.startswith("CSV") else "txt"),
            mime="text/csv" if fmt.startswith("CSV") else "text/plain",
            type="primary", use_container_width=True,
        )

# ─────────── TAB 4: IMPORTAR / EXPORTAR ───────────
with tab4:
    st.markdown("### Importar")
    if st.session_state["_ext_bytes"] is None:
        st.info(
            "Carrega um ficheiro Excel na barra lateral.\n\n"
            "**Tipos suportados:**\n"
            "- 🟪 **Excel INA** (multi-sheet por categoria) — ex: Base_de_dados_17_04_ATUALIZADO.xlsx\n"
            "- 🟧 **SIOE** (ExportResultadosPesquisa*.xlsx) — exportação manual do sioe.pt"
        )
    else:
        nm = st.session_state["_ext_name"]
        st.write("**Ficheiro:** " + nm)
        eh_sioe = "export" in nm.lower() and "pesquisa" in nm.lower()
        try:
            if eh_sioe:
                ent_novo, dir_novo, log = importar_sioe(st.session_state["_ext_bytes"])
                st.success(f"SIOE: {len(ent_novo):,} entidades · {len(dir_novo):,} dirigentes detectados.")
            else:
                ent_novo, dir_novo, log = importar_excel_ina(st.session_state["_ext_bytes"])
                st.success(f"INA: {len(ent_novo):,} entidades · {len(dir_novo):,} dirigentes detectados.")
            with st.expander("Detalhe do import"):
                st.json(log)
            st.markdown("#### Pré-visualização")
            c1, c2 = st.columns(2)
            with c1:
                st.caption(f"Entidades ({len(ent_novo):,})")
                st.dataframe(ent_novo.head(30), use_container_width=True, hide_index=True)
            with c2:
                st.caption(f"Dirigentes ({len(dir_novo):,})")
                st.dataframe(dir_novo.head(30), use_container_width=True, hide_index=True)
            st.markdown("---")
            st.markdown("#### Aplicar à base")
            col_a, col_b = st.columns([1, 1])
            with col_a:
                if st.button("✅ Fazer merge na base", type="primary", use_container_width=True):
                    ent_base = st.session_state["ent"] if st.session_state["ent"] is not None else pd.DataFrame(columns=COLS_ENTIDADE)
                    dir_base = st.session_state["dir"] if st.session_state["dir"] is not None else pd.DataFrame(columns=COLS_DIRIGENTE)
                    ent_merged, log_e = merge_entidades(ent_base, ent_novo)
                    dir_merged, log_d = merge_dirigentes(dir_base, dir_novo, ent_merged, ent_novo)
                    st.session_state["ent"] = ent_merged
                    st.session_state["dir"] = dir_merged
                    st.session_state["alterado_ent"] = True
                    st.session_state["alterado_dir"] = True
                    st.session_state["log_import"] = {"entidades": log_e, "dirigentes": log_d}
                    st.session_state["msg"] = ("ok", f"✅ Merge OK. Entidades: {log_e}. Dirigentes: {log_d}. Não esqueças de 'Guardar alterações' na sidebar.")
                    st.session_state["_ext_bytes"] = None
                    st.session_state["_ext_name"] = None
                    st.session_state["_ext_key"] = None
                    st.session_state["_file_key"] = None
                    st.rerun()
            with col_b:
                if st.button("✕ Descartar ficheiro", use_container_width=True):
                    st.session_state["_ext_bytes"] = None
                    st.session_state["_ext_name"] = None
                    st.session_state["_ext_key"] = None
                    st.session_state["_file_key"] = None
                    st.rerun()
        except Exception as e:
            st.error(f"Erro ao processar o ficheiro: {e}")
            st.exception(e)

    st.markdown("---")
    st.markdown("### Exportar")
    if len(ent) == 0:
        st.info("Base vazia — nada para exportar.")
    else:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            used = set()
            cats_map = cat_ordem_map()
            ent_ord = ent.copy()
            ent_ord["__ord"] = ent_ord["categoria_id"].map(cats_map).fillna(999)
            for cat_id, grupo in ent_ord.sort_values(["__ord", "designacao"]).groupby("categoria_id", sort=False):
                sheet = safe_sheet_name(cat_nome(cat_id), used)
                grupo_out = grupo.drop(columns=["__ord"], errors="ignore")
                grupo_out.to_excel(writer, sheet_name=sheet, index=False)
            if len(dir_):
                dir_.to_excel(writer, sheet_name=safe_sheet_name("Dirigentes", used), index=False)
        st.download_button(
            "⬇ Descarregar base completa (.xlsx)",
            data=buf.getvalue(),
            file_name=f"AP_Contactos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

# ─────────── TAB 5: REVISÃO (DRE) ───────────
with tab5:
    st.markdown("### Sugestões a partir do DRE")
    st.caption("Lê publicações recentes no DRE e propõe alterações a dirigentes. Aprovação manual.")
    col1, col2 = st.columns([1, 3])
    with col1:
        dias = st.number_input("Dias", min_value=1, max_value=180, value=30, step=1)
    with col2:
        if st.button("🔎 Buscar sugestões no DRE", use_container_width=True):
            with st.spinner("A consultar DRE..."):
                try:
                    sug = sugestoes_a_partir_de_dre(ent, dias=int(dias))
                    guardar_pendentes(sug)
                    st.success(f"{len(sug)} sugestões guardadas.")
                except Exception as e:
                    st.error(f"Erro no DRE: {e}")
    pendentes = carregar_pendentes()
    if not pendentes:
        st.info("Sem sugestões pendentes.")
    else:
        st.write(f"**{len(pendentes)}** sugestões pendentes")
        df_sug = pd.DataFrame(pendentes)
        st.dataframe(df_sug, use_container_width=True, hide_index=True)
        if st.button("🗑 Limpar sugestões pendentes"):
            guardar_pendentes([])
            st.rerun()

# ─────────── TAB 6: HISTÓRICO ───────────
with tab6:
    st.markdown("### Histórico de dirigentes")
    if len(dir_) == 0:
        st.info("Sem dirigentes registados.")
    else:
        hist = dir_[dir_["fim"].fillna("") != ""] if "fim" in dir_.columns else pd.DataFrame()
        if len(hist) == 0:
            st.info("Nenhum dirigente com saída registada.")
        else:
            joined = hist.merge(ent[["entity_id", "designacao"]], on="entity_id", how="left")
            st.write(f"**{len(joined):,}** registos históricos")
            st.dataframe(
                joined[["designacao", "cargo", "nome", "email", "inicio", "fim", "fonte"]].rename(
                    columns={"designacao": "Entidade", "cargo": "Cargo", "nome": "Nome",
                             "email": "Email", "inicio": "Início", "fim": "Fim", "fonte": "Fonte"}
                ),
                use_container_width=True,
                hide_index=True,
            )
