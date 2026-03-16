"""
AP Contactos v2.1 — INA
Persistência via GitHub. Categorias oficiais AP. Actualizações automáticas mensais.
Fontes: DRE, transparencia.gov.pt, dados.gov.pt, SIOE (manual)
"""
import base64
import html
import io
import json
import re
import unicodedata
from datetime import datetime

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="AP Contactos", page_icon="🏛️", layout="wide")

# ─────────────────────────────────────────────────────
# AUTENTICAÇÃO
# ─────────────────────────────────────────────────────
_PASSWORD = "INA#Contactos2026!"

def check_password():
    if st.session_state.get("autenticado"):
        return
    st.markdown("""
        <div style='max-width:400px;margin:80px auto 0;text-align:center'>
            <div style='font-size:3rem'>🏛️</div>
            <h2 style='color:#1F4E79;margin-bottom:4px'>AP Contactos</h2>
            <p style='color:#888;margin-bottom:28px;font-size:.95rem'>
                Instituto Nacional de Administração
            </p>
        </div>
    """, unsafe_allow_html=True)
    col = st.columns([1, 1, 1])[1]
    with col:
        pwd = st.text_input("", type="password", placeholder="Password de acesso",
                            label_visibility="collapsed")
        if st.button("Entrar →", use_container_width=True, type="primary"):
            if pwd == _PASSWORD:
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("Password incorrecta.")
    st.stop()

check_password()

# ─────────────────────────────────────────────────────
# CATEGORIAS OFICIAIS DA AP PORTUGUESA
# Baseadas na Lei n.º 4/2004 e estrutura do SIOE
# ─────────────────────────────────────────────────────

# Hierarquia oficial
CATEGORIAS_ORDEM = [
    "Administração Direta — Secretarias-Gerais",
    "Administração Direta — Direções-Gerais",
    "Administração Direta — Inspeções-Gerais",
    "Administração Direta — Gabinetes do Governo",
    "Administração Direta — Outros Serviços",
    "Administração Indireta — Institutos e Agências",
    "Administração Indireta — Fundos e Fundações",
    "Entidades Reguladoras Independentes",
    "Defesa e Segurança",
    "Regiões Autónomas — Secretarias Regionais",
    "Regiões Autónomas — Serviços Regionais",
    "Regiões Autónomas — Gabinetes",
    "Poder Local — Municípios",
    "Poder Local — Freguesias",
    "Poder Local — Entidades Intermunicipais",
    "Setor Empresarial do Estado",
    "Setor Empresarial Local",
    "Justiça — Tribunais",
    "Ensino Superior e Investigação",
    "Ensino Básico e Secundário",
    "Academias e Associações Científicas",
    "Cátedras UNESCO",
    "Fundações",
    "Associações",
    "Estruturas Consultivas e Temporárias",
    "Outros",
]

# Mapeamento SIOE → Categoria oficial
MAPA_SIOE = {
    "Secretaria-geral":                                        "Administração Direta — Secretarias-Gerais",
    "Direção-geral":                                           "Administração Direta — Direções-Gerais",
    "Inspeção-geral":                                          "Administração Direta — Inspeções-Gerais",
    "Inspeção Regional":                                       "Administração Direta — Inspeções-Gerais",
    "Gabinete 1.º Ministro":                                   "Administração Direta — Gabinetes do Governo",
    "Gabinete Ministro":                                       "Administração Direta — Gabinetes do Governo",
    "Gabinete Secretário de Estado":                           "Administração Direta — Gabinetes do Governo",
    "Gabinete":                                                "Administração Direta — Gabinetes do Governo",
    "Serviço de Apoio":                                        "Administração Direta — Outros Serviços",
    "Instituto Público":                                       "Administração Indireta — Institutos e Agências",
    "Agência":                                                 "Administração Indireta — Institutos e Agências",
    "Centro de Formação Profissional":                         "Administração Indireta — Institutos e Agências",
    "Banco Central":                                           "Administração Indireta — Institutos e Agências",
    "Fundo Autónomo":                                          "Administração Indireta — Fundos e Fundações",
    "Fundo da Segurança Social":                               "Administração Indireta — Fundos e Fundações",
    "Entidade Administrativa Independente":                    "Entidades Reguladoras Independentes",
    "Órgão Independente":                                      "Entidades Reguladoras Independentes",
    "Força de Segurança":                                      "Defesa e Segurança",
    "Forças Armadas":                                          "Defesa e Segurança",
    "Gabinete Secretário Regional":                            "Regiões Autónomas — Gabinetes",
    "Gabinete Presidente Regional":                            "Regiões Autónomas — Gabinetes",
    "Gabinete Vice-Presidente Regional":                       "Regiões Autónomas — Gabinetes",
    "Gabinete do Representante da República":                  "Regiões Autónomas — Gabinetes",
    "Direção Regional":                                        "Regiões Autónomas — Serviços Regionais",
    "Município":                                               "Poder Local — Municípios",
    "Serviço Municipalizado e Intermunicipalizado":            "Poder Local — Municípios",
    "Junta de Freguesia":                                      "Poder Local — Freguesias",
    "Associação de Municípios de fins específicos":            "Poder Local — Entidades Intermunicipais",
    "Comunidade intermunicipal":                               "Poder Local — Entidades Intermunicipais",
    "Federação de Municípios":                                 "Poder Local — Entidades Intermunicipais",
    "Área Metropolitana":                                      "Poder Local — Entidades Intermunicipais",
    "Associação de Freguesias":                                "Poder Local — Entidades Intermunicipais",
    "Entidade Pública Empresarial":                            "Setor Empresarial do Estado",
    "Entidade Pública Empresarial Regional":                   "Setor Empresarial do Estado",
    "Entidade Empresarial Regional":                           "Setor Empresarial do Estado",
    "Sociedade Anónima":                                       "Setor Empresarial do Estado",
    "Sociedade por Quotas":                                    "Setor Empresarial do Estado",
    "Cooperativa":                                             "Setor Empresarial do Estado",
    "Agrupamento Complementar de Empresas ":                   "Setor Empresarial do Estado",
    "Empresa Municipal":                                       "Setor Empresarial Local",
    "Empresa Intermunicipal":                                  "Setor Empresarial Local",
    "Entidade Empresarial Municipal":                          "Setor Empresarial Local",
    "Tribunal":                                                "Justiça — Tribunais",
    "Unidade Orgânica de Ensino e Investigação":               "Ensino Superior e Investigação",
    "Unidade Orgânica de Investigação":                        "Ensino Superior e Investigação",
    "Estabelecimento de educação e ensino básico e secundário":"Ensino Básico e Secundário",
    "Associação":                                              "Associações",
    "Fundação":                                                "Fundações",
    "Estrutura temporária - comissão":                         "Estruturas Consultivas e Temporárias",
    "Estrutura temporária - estrutura de missão":              "Estruturas Consultivas e Temporárias",
    "Estrutura temporária - grupo de projeto":                 "Estruturas Consultivas e Temporárias",
    "Estrutura temporária - grupo de trabalho":                "Estruturas Consultivas e Temporárias",
    "Órgão consultivo":                                        "Estruturas Consultivas e Temporárias",
    "Entidade Regional de Turismo":                            "Outros",
    "Estrutura atípica":                                       "Outros",
}

# Mapeamento das sheets INA → Categoria oficial
MAPA_INA = {
    "Secretaria-Geral Governo":  "Administração Direta — Secretarias-Gerais",
    "Secretaria-Geral Governo ": "Administração Direta — Secretarias-Gerais",
    "Secretarias-Gerais":        "Administração Direta — Secretarias-Gerais",
    "Direções-Gerais":           "Administração Direta — Direções-Gerais",
    "Inspeções-Gerais":          "Administração Direta — Inspeções-Gerais",
    "Autoridades":               "Entidades Reguladoras Independentes",
    "Institutos-Agências":       "Administração Indireta — Institutos e Agências",
    "Direções-Regionais":        "Regiões Autónomas — Serviços Regionais",
    "Empresas Públicas":         "Setor Empresarial do Estado",
    "Fundações":                 "Fundações",
    "Comissões":                 "Estruturas Consultivas e Temporárias",
    "Academias":                 "Academias e Associações Científicas",
    "Cátedras Nacionais":        "Cátedras UNESCO",
}

# Sheets a ignorar
SHEETS_IGNORAR = {"Menu", "Conselho Editorial", "Conselho Estratégico"}

# Configuração das sheets INA
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

COLS_BASE = ["sigla_entidade","designacao","ministerio","tipo_entidade",
             "orgao_direcao","nome_dirigente","email","contacto","categoria","fonte","id"]

GITHUB_DATA_PATH = "dados/contactos.csv"
GITHUB_LOG_PATH  = "dados/log_atualizacoes.json"


# ─────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────
def limpar(v):
    if not isinstance(v, str):
        v = "" if v is None or (isinstance(v, float) and str(v) == "nan") else str(v)
    return "".join(c for c in v if unicodedata.category(c) not in ("Cf","Cc") or c in ("\n","\t")).strip()

def email_ok(e):
    e = limpar(str(e)).lower()
    if not e or e in ("nan","none","n/d"): return ""
    return e if re.match(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$", e) else ""

def garantir_cols(df):
    for col in COLS_BASE:
        if col not in df.columns: df[col] = ""
    return df

def categoria_para_ordem(cat):
    try: return CATEGORIAS_ORDEM.index(cat)
    except ValueError: return 999


# ─────────────────────────────────────────────────────
# GITHUB — CARREGAR E GUARDAR
# ─────────────────────────────────────────────────────
def _github_headers():
    token = st.secrets.get("GITHUB_TOKEN", "")
    return {"Authorization": f"token {token}", "Content-Type": "application/json"} if token else {}

def _github_repo():
    return st.secrets.get("GITHUB_REPO", "")

def github_configurado():
    return bool(st.secrets.get("GITHUB_TOKEN","")) and bool(st.secrets.get("GITHUB_REPO",""))

def carregar_do_github():
    """Tenta carregar contactos.csv do GitHub. Devolve (df, sha) ou (None, None)."""
    repo = _github_repo()
    if not repo: return None, None
    try:
        url = f"https://api.github.com/repos/{repo}/contents/{GITHUB_DATA_PATH}"
        r = requests.get(url, headers=_github_headers(), timeout=15)
        if r.status_code == 404: return None, None
        if r.status_code != 200: return None, None
        data = r.json()
        content = base64.b64decode(data["content"]).decode("utf-8")
        sha = data["sha"]
        df = pd.read_csv(io.StringIO(content), dtype=str).fillna("")
        df = garantir_cols(df)
        if "id" not in df.columns or df["id"].eq("").all():
            df = df.reset_index(drop=True)
            df["id"] = df.index
        df["id"] = pd.to_numeric(df["id"], errors="coerce").fillna(0).astype(int)
        return df, sha
    except Exception:
        return None, None

def guardar_no_github(df, sha=None):
    """Guarda df como CSV no GitHub. Devolve (sucesso, mensagem, novo_sha)."""
    repo = _github_repo()
    if not repo: return False, "GitHub não configurado", sha
    try:
        cols = [c for c in COLS_BASE if c in df.columns]
        csv_str = df[cols].to_csv(index=False)
        content_b64 = base64.b64encode(csv_str.encode("utf-8")).decode()
        payload = {
            "message": f"AP Contactos — actualização automática {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "content": content_b64,
        }
        if sha: payload["sha"] = sha
        url = f"https://api.github.com/repos/{repo}/contents/{GITHUB_DATA_PATH}"
        r = requests.put(url, headers=_github_headers(), data=json.dumps(payload), timeout=30)
        if r.status_code in (200, 201):
            novo_sha = r.json().get("content", {}).get("sha", sha)
            n = len(df)
            n_em = (df["email"].str.len() > 3).sum()
            return True, f"✅ Guardado com sucesso! {n:,} registos ({n_em:,} emails) guardados na base permanente.", novo_sha
        else:
            return False, f"Erro GitHub {r.status_code}: {r.text[:200]}", sha
    except Exception as e:
        return False, f"Erro ao guardar: {e}", sha


# ─────────────────────────────────────────────────────
# CARREGAR LOG DE ACTUALIZAÇÕES
# ─────────────────────────────────────────────────────
def carregar_log_github():
    """Lê o log de actualizações automáticas do GitHub."""
    try:
        url = f"https://api.github.com/repos/{_github_repo()}/contents/{GITHUB_LOG_PATH}"
        r = requests.get(url, headers=_github_headers(), timeout=10)
        if r.status_code != 200: return None
        content = base64.b64decode(r.json()["content"]).decode("utf-8")
        return json.loads(content)
    except Exception:
        return None


# ─────────────────────────────────────────────────────
# CARREGAR BASE PRINCIPAL (Excel INA)
# ─────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def carregar_base_excel(file_bytes: bytes) -> pd.DataFrame:
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    frames = []
    for sheet in xl.sheet_names:
        ss = sheet.strip()
        if ss in SHEETS_IGNORAR: continue
        cfg = SHEET_CONFIG.get(sheet) or SHEET_CONFIG.get(ss)
        if cfg is None: continue
        try:
            df = xl.parse(sheet, header=cfg["header_row"])
            expected = cfg["cols"]
            df = df.iloc[:, :len(expected)]
            while len(df.columns) < len(expected): df[f"_x{len(df.columns)}"] = ""
            df.columns = expected
            # Categoria oficial
            df["categoria"] = MAPA_INA.get(ss, MAPA_INA.get(sheet, "Outros"))
            df["fonte"] = "base_principal"
            frames.append(df)
        except Exception:
            pass
    if not frames: return pd.DataFrame(columns=COLS_BASE)
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

    email_s = pd.Series("", index=df.index)
    tel_s   = pd.Series("", index=df.index)
    for tc, cc in zip(tipo_cols, contacto_cols):
        t = df[tc].fillna("").astype(str).str.strip().str.lower()
        v = df[cc].fillna("").astype(str).str.strip()
        email_s = email_s.where(~((t == "email")   & (email_s == "")), v.str.lower())
        tel_s   = tel_s.where(~((t == "telefone")  & (tel_s   == "")), v)
    email_s = email_s.where(
        email_s.str.match(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$"), ""
    )

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

    def gc(n): return df.get(n, pd.Series("", index=df.index)).fillna("").astype(str).str.strip()

    tipo_serie = gc("Tipo de Entidade")
    novo = pd.DataFrame({
        "designacao":     gc("Designação"),
        "sigla_entidade": gc("Sigla"),
        "ministerio":     gc("Ministério/Secretaria Regional"),
        "tipo_entidade":  tipo_serie,
        "email":          email_s,
        "contacto":       tel_s,
        "nome_dirigente": nome_s,
        "orgao_direcao":  cargo_s,
        "categoria":      tipo_serie.map(MAPA_SIOE).fillna("Outros"),
        "fonte":          "SIOE",
    })
    return novo[novo["designacao"].str.len() > 2].copy()


# ─────────────────────────────────────────────────────
# FUNDIR COM BASE
# ─────────────────────────────────────────────────────
def fundir_com_base(df_novo: pd.DataFrame) -> tuple:
    df_novo = garantir_cols(df_novo.copy())
    emails_base = set(st.session_state["df"]["email"].str.lower())
    emails_base.discard("")
    df_novo = df_novo[~df_novo["email"].isin(emails_base) | (df_novo["email"] == "")].copy()
    if df_novo.empty: return 0, 0
    id_max = int(st.session_state["df"]["id"].max()) + 1
    df_novo = df_novo.reset_index(drop=True)
    df_novo["id"] = df_novo.index + id_max
    n_antes = len(st.session_state["df"])
    st.session_state["df"] = pd.concat(
        [st.session_state["df"], df_novo[COLS_BASE]], ignore_index=True
    )
    st.session_state["dados_alterados"] = True
    return len(st.session_state["df"]) - n_antes, int((df_novo["email"].str.len() > 3).sum())


def atualizar_com_sioe(df_novo: pd.DataFrame) -> tuple:
    df_novo = garantir_cols(df_novo.copy())
    df_base = st.session_state["df"].copy()
    campos  = ["nome_dirigente","orgao_direcao","contacto","ministerio","sigla_entidade","tipo_entidade"]
    emails_novos = set(df_novo[df_novo["email"].str.len()>3]["email"].str.lower())
    emails_base  = set(df_base[df_base["email"].str.len()>3]["email"].str.lower())
    comuns    = emails_novos & emails_base
    so_novos  = emails_novos - emails_base
    n_upd = 0
    for email in comuns:
        mb = df_base["email"].str.lower() == email
        mn = df_novo["email"].str.lower() == email
        if not mn.any(): continue
        row = df_novo[mn].iloc[0]
        for c in campos:
            v = str(row.get(c,"")).strip()
            if v and v not in ("nan","None",""): df_base.loc[mb, c] = v
        df_base.loc[mb, "fonte"] = "SIOE (actualizado)"
        n_upd += 1
    df_so = df_novo[df_novo["email"].isin(so_novos) | (df_novo["email"] == "")]
    n_add = 0
    if not df_so.empty:
        id_max = int(df_base["id"].max()) + 1
        df_so = df_so.reset_index(drop=True).copy()
        df_so["id"] = df_so.index + id_max
        df_base = pd.concat([df_base, df_so[COLS_BASE]], ignore_index=True)
        n_add = len(df_so)
    st.session_state["df"] = df_base
    st.session_state["dados_alterados"] = True
    return n_upd, n_add


# ─────────────────────────────────────────────────────
# EXPORTAR EXCEL
# ─────────────────────────────────────────────────────
def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    cols = ["categoria","sigla_entidade","designacao","ministerio","tipo_entidade",
            "orgao_direcao","nome_dirigente","email","contacto","fonte"]
    cols_ok = [c for c in cols if c in df.columns]
    buf = io.BytesIO()
    cats_ordenadas = sorted(df["categoria"].dropna().unique(), key=categoria_para_ordem)
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df[cols_ok].to_excel(writer, sheet_name="Todos", index=False)
        for cat in cats_ordenadas:
            df[df["categoria"]==cat][cols_ok].to_excel(
                writer, sheet_name=str(cat)[:31], index=False)
    return buf.getvalue()


# ─────────────────────────────────────────────────────
# ESTADO INICIAL
# ─────────────────────────────────────────────────────
defaults = {
    "df": None, "sel": {}, "github_sha": None,
    "dados_alterados": False, "import_msg": None,
    "_file_key": None, "_ext_key": None, "_ext_bytes": None, "_ext_name": None,
}
for k, v in defaults.items():
    if k not in st.session_state: st.session_state[k] = v


# ─────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────
st.markdown("""<style>
.metric-box{background:white;border-radius:10px;padding:14px 18px;
  border-left:4px solid #1a56db;margin-bottom:10px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.metric-box .n{font-size:1.9rem;font-weight:700;color:#1a56db}
.metric-box .l{font-size:.82rem;color:#666}
.email-box{background:#f8f9fa;border:1px solid #dee2e6;border-radius:8px;padding:12px 16px;
  font-family:monospace;font-size:.84rem;max-height:300px;overflow-y:auto;
  white-space:pre-wrap;word-break:break-all}
.save-bar{background:#fff3cd;border:1px solid #ffc107;border-radius:8px;
  padding:10px 16px;margin:8px 0;display:flex;align-items:center;gap:10px}
</style>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏛️ AP Contactos")

    # ── Estado da ligação ao GitHub ──
    if github_configurado():
        st.success("🔗 GitHub configurado")
        if st.session_state["df"] is None:
            with st.spinner("A carregar dados do GitHub..."):
                df_gh, sha = carregar_do_github()
                if df_gh is not None:
                    st.session_state["df"]         = df_gh
                    st.session_state["github_sha"] = sha
                    st.session_state["dados_alterados"] = False
    else:
        st.warning("⚠️ GitHub não configurado\n\nAdiciona `GITHUB_TOKEN` e `GITHUB_REPO` nos Secrets do Streamlit para persistência automática.")

    st.markdown("---")
    st.markdown("### 📂 Carregar base Excel")
    st.caption("Usa para carregar a base INA pela primeira vez, ou substituir os dados actuais.")
    uploaded = st.file_uploader("Base INA (.xlsx)", type=["xlsx"], label_visibility="collapsed")

    if uploaded is not None:
        file_bytes = uploaded.getvalue()
        file_key   = f"{uploaded.name}_{len(file_bytes)}"
        if st.session_state["_file_key"] != file_key:
            st.session_state["_file_key"]  = file_key
            st.session_state["sel"]        = {}
            st.session_state["import_msg"] = None
            st.session_state["_ext_key"]   = None
            st.cache_data.clear()
            with st.spinner("A carregar base Excel..."):
                st.session_state["df"] = carregar_base_excel(file_bytes)
                st.session_state["dados_alterados"] = True

    if st.session_state["df"] is not None:
        n_tot  = len(st.session_state["df"])
        n_mail = (st.session_state["df"]["email"].str.len() > 3).sum()
        st.success(f"✅ {n_tot:,} registos  |  {n_mail:,} emails")
        if st.session_state["dados_alterados"] and github_configurado():
            if st.button("💾 Guardar na base permanente", type="primary", use_container_width=True):
                with st.spinner("A guardar no GitHub..."):
                    ok, msg, novo_sha = guardar_no_github(
                        st.session_state["df"], st.session_state["github_sha"]
                    )
                if ok:
                    st.session_state["github_sha"]    = novo_sha
                    st.session_state["dados_alterados"] = False
                    st.session_state["import_msg"]    = {"texto": msg}
                    st.rerun()
                else:
                    st.error(msg)
    else:
        st.info("Nenhum dado carregado.")

    st.markdown("---")
    st.markdown("### 🔍 Filtros")


# ─────────────────────────────────────────────────────
# GUARD
# ─────────────────────────────────────────────────────
if st.session_state["df"] is None:
    st.title("🏛️ Contactos — AP Portuguesa")
    if github_configurado():
        st.info("A tentar carregar dados do GitHub… Se demorar, carrega o ficheiro Excel manualmente na barra lateral.")
    else:
        st.info("👆 Carrega o ficheiro Excel principal na barra lateral para começar.")
        st.markdown("""
        **Para activar persistência automática**, adiciona nos Secrets do Streamlit Cloud:
        ```toml
        GITHUB_TOKEN = "ghp_o_teu_token"
        GITHUB_REPO  = "bernardohille-cmyk/Pesquisar-emails-FP"
        ```
        Depois de guardar uma vez, os dados ficam sempre disponíveis.
        """)
    st.stop()

df = st.session_state["df"]


# ─────────────────────────────────────────────────────
# FILTROS SIDEBAR
# ─────────────────────────────────────────────────────
with st.sidebar:
    pesquisa = st.text_input("🔎 Pesquisar", "")
    cats_disp = sorted(df["categoria"].dropna().unique(), key=categoria_para_ordem)
    cats_sel  = st.multiselect("Categoria (oficial)", cats_disp)
    _prefixos = ("Ministério","Secretaria","Presidência","Vice-Presidência")
    min_disp  = sorted(m for m in df["ministerio"].fillna("").unique()
                       if m and any(m.startswith(p) for p in _prefixos))
    min_sel   = st.multiselect("Ministério / Tutela", min_disp)
    so_email  = st.checkbox("Só com email", value=False)
    st.markdown("---")
    if st.button("🗑️ Limpar & recarregar", use_container_width=True):
        for k in list(st.session_state.keys()): del st.session_state[k]
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
if cats_sel:  df_f = df_f[df_f["categoria"].isin(cats_sel)]
if min_sel:   df_f = df_f[df_f["ministerio"].isin(min_sel)]
if so_email:  df_f = df_f[df_f["email"].str.len() > 3]
tem_filtro = bool(pesquisa or cats_sel or min_sel or so_email)


# ─────────────────────────────────────────────────────
# CABEÇALHO
# ─────────────────────────────────────────────────────
st.title("🏛️ Contactos — AP Portuguesa")

if st.session_state.get("import_msg"):
    st.success(st.session_state["import_msg"]["texto"])
    if st.button("✖ Fechar", key="fechar_msg"):
        st.session_state["import_msg"] = None
        st.rerun()

if st.session_state["dados_alterados"] and github_configurado():
    st.warning("⚠️ Tens alterações não guardadas. Clica **💾 Guardar na base permanente** na barra lateral.", icon="⚠️")

ev = df_f[df_f["email"].str.len() > 3]
for col, num, lbl in zip(
    st.columns(4),
    [f"{len(df_f):,}", f"{len(ev):,}", f"{df_f['designacao'].nunique():,}",
     f"{df_f['ministerio'].replace('',pd.NA).dropna().nunique():,}"],
    ["Registos", "Com email", "Entidades", "Ministérios"],
):
    col.markdown(f'<div class="metric-box"><div class="n">{num}</div><div class="l">{lbl}</div></div>',
                 unsafe_allow_html=True)


# ─────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 Tabela", "☑️ Selecionar & Copiar Emails",
    "➕ Importar / Actualizar", "📥 Descarregar",
    "🤖 Actualizações Automáticas",
])


# ══════════════════════════════════════════════════════
# TAB 1 — TABELA
# ══════════════════════════════════════════════════════
with tab1:
    lbl_map = {"categoria":"Categoria (oficial)","sigla_entidade":"Sigla","designacao":"Entidade",
               "ministerio":"Ministério","orgao_direcao":"Órgão/Cargo","nome_dirigente":"Dirigente",
               "email":"Email","contacto":"Contacto","fonte":"Fonte"}
    cols_v = [c for c in lbl_map if c in df_f.columns]
    st.dataframe(df_f[cols_v].rename(columns=lbl_map),
                 use_container_width=True, height=520, hide_index=True)
    st.caption(f"A mostrar {len(df_f):,} de {len(df):,} registos totais")

    df_em_f = df_f[df_f["email"].str.len() > 3].drop_duplicates(subset=["email"])
    n_em_f  = len(df_em_f)
    if n_em_f > 0:
        st.markdown("---")
        st.markdown(f"**📧 Exportação rápida — {n_em_f:,} emails no filtro actual:**")
        eq1, eq2, eq3 = st.columns(3)
        with eq1:
            st.download_button(f"⬇️ BCC Outlook ({n_em_f:,})",
                data=("; ".join(sorted(df_em_f["email"].tolist()))).encode("utf-8"),
                file_name=f"emails_BCC_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain", use_container_width=True, type="primary",
                help="Separados por ; — cola no campo Cco do Outlook")
        with eq2:
            st.download_button(f"⬇️ Um por linha ({n_em_f:,})",
                data=("\n".join(sorted(df_em_f["email"].tolist()))).encode("utf-8"),
                file_name=f"emails_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain", use_container_width=True)
        with eq3:
            csv_em = df_em_f[["categoria","designacao","nome_dirigente","email"]].to_csv(
                index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(f"⬇️ CSV com nomes ({n_em_f:,})",
                data=csv_em,
                file_name=f"emails_nomes_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv", use_container_width=True)
    elif len(df_f) > 0:
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
        c1b, c2b, c3b = st.columns([1,1,2])
        with c1b:
            if st.button("✅ Todos", use_container_width=True):
                for i in df_ce["id"].tolist(): st.session_state["sel"][i] = True
                st.rerun()
        with c2b:
            if st.button("❌ Limpar", use_container_width=True):
                st.session_state["sel"] = {}
                st.rerun()
        with c3b:
            cat_r = st.selectbox("Adicionar categoria:", ["—"] + sorted(df_ce["categoria"].unique(),
                                 key=categoria_para_ordem), key="cat_r")
            if cat_r != "—" and st.button(f"➕ '{cat_r}'", use_container_width=True):
                for i in df_ce[df_ce["categoria"]==cat_r]["id"].tolist(): st.session_state["sel"][i] = True
                st.rerun()

        st.markdown("---")
        for cat in sorted(df_ce["categoria"].unique(), key=categoria_para_ordem):
            sub = df_ce[df_ce["categoria"]==cat]
            n_s = sum(1 for i in sub["id"] if st.session_state["sel"].get(i, False))
            with st.expander(f"**{cat}** — {len(sub)} com email  |  ✓ {n_s} seleccionados"):
                sc1, sc2 = st.columns(2)
                with sc1:
                    if st.button(f"✅ Todos", key=f"sa_{cat}", use_container_width=True):
                        for i in sub["id"].tolist(): st.session_state["sel"][i] = True
                        st.rerun()
                with sc2:
                    if st.button(f"❌ Nenhum", key=f"sn_{cat}", use_container_width=True):
                        for i in sub["id"].tolist(): st.session_state["sel"][i] = False
                        st.rerun()
                for _, row in sub.iterrows():
                    idx = row["id"]
                    lbl_chk = f"{row['designacao']}  —  `{row['email']}`"
                    if str(row.get("nome_dirigente","")).strip(): lbl_chk += f"  ({row['nome_dirigente']})"
                    v = st.checkbox(lbl_chk, value=st.session_state["sel"].get(idx,False), key=f"chk_{idx}")
                    if v != st.session_state["sel"].get(idx, False): st.session_state["sel"][idx] = v

        st.markdown("---")
        ids_sel = [i for i,v in st.session_state["sel"].items() if v]
        df_sel  = df_ce[df_ce["id"].isin(ids_sel)]
        st.markdown(f"### 📋 {len(df_sel)} seleccionados")
        if not df_sel.empty:
            cf1, cf2 = st.columns([2,1])
            with cf1: fmt = st.radio("Formato", ["Um por linha","Separados por ; (BCC)","CSV"], horizontal=True)
            with cf2:
                inc_nome = st.checkbox("Incluir nome", value=False)
                dedup    = st.checkbox("Sem duplicados", value=True)
            df_exp = df_sel.drop_duplicates(subset=["email"]) if dedup else df_sel
            lista  = sorted(df_exp["email"].tolist())
            if inc_nome:
                pares = df_exp[["nome_dirigente","email"]]
                if fmt == "Um por linha":           texto = "\n".join(f"{r['nome_dirigente']} <{r['email']}>" for _,r in pares.iterrows())
                elif fmt == "Separados por ; (BCC)":texto = "; ".join(f"{r['nome_dirigente']} <{r['email']}>" for _,r in pares.iterrows())
                else:                               texto = "Nome,Email\n" + "\n".join(f'"{r["nome_dirigente"]}",{r["email"]}' for _,r in pares.iterrows())
            else:
                if fmt == "Um por linha":           texto = "\n".join(lista)
                elif fmt == "Separados por ; (BCC)":texto = "; ".join(lista)
                else:                               texto = "Email\n" + "\n".join(lista)
            st.markdown(f"**{len(lista)} emails:**")
            st.markdown(f'<div class="email-box">{html.escape(texto)}</div>', unsafe_allow_html=True)
            st.download_button("⬇️ Descarregar (.txt)", data=texto.encode("utf-8"),
                file_name=f"emails_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain", use_container_width=True)
        else:
            st.info("Nenhum organismo seleccionado.")


# ══════════════════════════════════════════════════════
# TAB 3 — IMPORTAR / ACTUALIZAR
# ══════════════════════════════════════════════════════
with tab3:
    st.markdown("### ➕ Importar / Actualizar dados")

    uploaded_ext = st.file_uploader("Carregar ficheiro externo (.xlsx)", type=["xlsx"],
                                    key="uploader_ext", label_visibility="collapsed")
    if uploaded_ext is not None:
        ext_key = f"{uploaded_ext.name}_{uploaded_ext.size}"
        if st.session_state["_ext_key"] != ext_key:
            st.session_state["_ext_key"]   = ext_key
            st.session_state["_ext_bytes"] = uploaded_ext.getvalue()
            st.session_state["_ext_name"]  = uploaded_ext.name

    ext_bytes = st.session_state["_ext_bytes"]
    ext_name  = st.session_state["_ext_name"] or ""

    if ext_bytes is None:
        st.info("Suporta:\n- 🏛️ **Exportação SIOE** (`ExportResultadosPesquisa*.xlsx`) — automático, com categorias oficiais\n- 📋 **Qualquer outro Excel** — mapeamento manual de colunas")
    else:
        parece_sioe = "export" in ext_name.lower() and "pesquisa" in ext_name.lower()
        if not parece_sioe:
            try:
                _t = pd.read_excel(io.BytesIO(ext_bytes), header=0, nrows=1)
                _t.columns = [str(c).strip() for c in _t.columns]
                parece_sioe = "Designação" in _t.columns and "Código SIOE" in _t.columns
            except Exception: pass

        st.markdown(f"**Ficheiro:** `{ext_name}`")
        if st.button("🗑️ Remover", key="rm_ext"):
            st.session_state["_ext_key"] = st.session_state["_ext_bytes"] = st.session_state["_ext_name"] = None
            st.rerun()
        st.markdown("---")

        # ── SIOE ──
        if parece_sioe:
            st.success("✅ Ficheiro SIOE detectado — categorias oficiais atribuídas automaticamente")
            try: n_lin = len(pd.read_excel(io.BytesIO(ext_bytes), header=0, usecols=[0]))
            except: n_lin = "?"
            cm, ci = st.columns([1,2])
            with cm:
                st.metric("Entidades", f"{n_lin:,}" if isinstance(n_lin,int) else n_lin)
                modo_atualizacao = st.toggle("🔄 Modo actualização", value=False,
                    help="ON: sobrescreve dirigente/cargo/telefone dos registos existentes com dados SIOE mais recentes.\nOFF: só acrescenta novos emails.")
                btn_sioe = st.button("🚀 Importar SIOE", type="primary", use_container_width=True)
            with ci:
                if modo_atualizacao:
                    st.warning("**Modo actualização** — sobrescreve campos dos registos existentes. Emails não mudam. Registos não são apagados.", icon="🔄")
                else:
                    st.info("Modo normal — só acrescenta emails novos, não toca nos existentes.")

            if btn_sioe:
                with st.spinner("A processar SIOE…"):
                    try:
                        df_novo = importar_sioe(ext_bytes)
                        if df_novo.empty:
                            st.error("Nenhum registo válido.")
                        elif modo_atualizacao:
                            n_upd, n_add = atualizar_com_sioe(df_novo)
                            st.session_state["import_msg"] = {"texto":
                                f"🔄 **{n_upd:,}** registos actualizados, **{n_add:,}** novos acrescentados. "
                                f"Base: **{len(st.session_state['df']):,}** registos. "
                                f"{'💾 Não te esqueças de guardar!' if github_configurado() else ''}"}
                            st.rerun()
                        else:
                            n_add, n_em = fundir_com_base(df_novo)
                            if n_add == 0:
                                st.warning("Todos os emails do SIOE já existem na base.")
                            else:
                                st.session_state["import_msg"] = {"texto":
                                    f"✅ **{n_add:,}** registos adicionados, **{n_em:,}** com email. "
                                    f"Base: **{len(st.session_state['df']):,}** registos. "
                                    f"{'💾 Não te esqueças de guardar!' if github_configurado() else ''}"}
                                st.rerun()
                    except Exception as e:
                        st.error(f"Erro: {e}")

        # ── Manual ──
        st.markdown("---")
        with st.expander("🔧 Mapeamento manual de colunas", expanded=not parece_sioe):
            try:
                xl_ext = pd.ExcelFile(io.BytesIO(ext_bytes))
                cs1, cs2 = st.columns(2)
                with cs1: sheet_esc = st.selectbox("Sheet", xl_ext.sheet_names, key="m_sheet")
                with cs2: hrow = int(st.number_input("Linha cabeçalho (0=primeira)", 0, 10, 0, key="m_hrow"))
                df_prev = xl_ext.parse(sheet_esc, header=hrow, nrows=3)
                df_prev.columns = [str(c).strip() for c in df_prev.columns]
                st.dataframe(df_prev, use_container_width=True, hide_index=True)
                opcoes = ["— (não mapear)"] + list(df_prev.columns)
                campos_map = {"designacao":"Nome entidade *(obrigatório)*","email":"Email *(obrigatório)*",
                              "nome_dirigente":"Nome dirigente","orgao_direcao":"Cargo/Órgão",
                              "ministerio":"Ministério","sigla_entidade":"Sigla",
                              "tipo_entidade":"Tipo entidade","contacto":"Telefone"}
                mapa = {}
                cm1, cm2 = st.columns(2)
                for k, (campo, desc) in enumerate(campos_map.items()):
                    with (cm1 if k%2==0 else cm2):
                        esc = st.selectbox(desc, opcoes, key=f"mc_{campo}")
                        if esc != "— (não mapear)": mapa[campo] = esc

                cat_nome = st.selectbox("Categoria oficial", CATEGORIAS_ORDEM, key="m_cat")

                if st.button("🔀 Fundir com a base", type="primary", use_container_width=True, key="btn_manual"):
                    if "designacao" not in mapa or "email" not in mapa:
                        st.error("Tens de mapear pelo menos Nome e Email.")
                    else:
                        try:
                            df_m = xl_ext.parse(sheet_esc, header=hrow)
                            df_m.columns = [str(c).strip() for c in df_m.columns]
                            novo = pd.DataFrame()
                            for campo in ["designacao","sigla_entidade","ministerio","tipo_entidade","orgao_direcao","nome_dirigente","email","contacto"]:
                                col_e = mapa.get(campo,"")
                                novo[campo] = df_m[col_e].apply(limpar) if col_e and col_e in df_m.columns else ""
                            novo["email"]     = novo["email"].apply(email_ok)
                            novo["categoria"] = cat_nome
                            novo["fonte"]     = f"importado:{ext_name}"
                            novo = novo[novo["designacao"].str.len() > 2].copy()
                            if novo.empty:
                                st.warning("Nenhum registo válido.")
                            else:
                                n_add, n_em = fundir_com_base(novo)
                                if n_add == 0:
                                    st.warning("Todos os registos já existem na base.")
                                else:
                                    st.session_state["import_msg"] = {"texto":
                                        f"✅ **{n_add:,}** registos adicionados à categoria **{cat_nome}** "
                                        f"(**{n_em:,}** com email). "
                                        f"{'💾 Não te esqueças de guardar!' if github_configurado() else ''}"}
                                    st.rerun()
                        except Exception as e:
                            st.error(f"Erro: {e}")
            except Exception as e:
                st.error(f"Erro: {e}")


# ══════════════════════════════════════════════════════
# TAB 4 — DESCARREGAR
# ══════════════════════════════════════════════════════
with tab4:
    st.markdown("### 📥 Descarregar base de dados")
    df_dl = st.session_state["df"]
    n_tot = len(df_dl); n_mail = (df_dl["email"].str.len()>3).sum()
    st.info(f"Base actual: **{n_tot:,}** registos  |  **{n_mail:,}** com email  |  **{n_tot-n_mail:,}** sem email")

    cd1, cd2, cd3 = st.columns(3)
    with cd1:
        st.markdown("**Excel completo** (uma sheet por categoria oficial)")
        st.download_button("⬇️ Excel completo",
            data=df_to_excel_bytes(df_dl),
            file_name=f"AP_Contactos_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary")
    with cd2:
        st.markdown(f"**Excel filtrado** ({len(df_f):,} registos)")
        if not tem_filtro:
            st.warning("Activa um filtro primeiro.", icon="⚠️")
        else:
            st.download_button(f"⬇️ Excel filtrado ({len(df_f):,})",
                data=df_to_excel_bytes(df_f.drop(columns=["id"],errors="ignore")),
                file_name=f"AP_Contactos_filtrado_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary")
    with cd3:
        st.markdown(f"**CSV filtrado** ({len(df_f):,} registos)")
        if not tem_filtro:
            st.warning("Activa um filtro primeiro.", icon="⚠️")
        else:
            csv_b = (df_f.drop(columns=["id"],errors="ignore")
                     .to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"))
            st.download_button(f"⬇️ CSV filtrado ({len(df_f):,})",
                data=csv_b,
                file_name=f"AP_Contactos_filtrado_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv", use_container_width=True)

    st.markdown("---")
    st.markdown("**Cobertura por categoria oficial:**")
    res = (df_dl.groupby("categoria")
           .agg(Total=("id","count"), Com_Email=("email", lambda x:(x.str.len()>3).sum()))
           .reset_index())
    res.columns = ["Categoria","Total","Com Email"]
    res["Sem Email"] = res["Total"] - res["Com Email"]
    res["Cobertura"] = (res["Com Email"]/res["Total"]*100).round(1).astype(str)+"%"
    res["_ord"] = res["Categoria"].apply(categoria_para_ordem)
    st.dataframe(res.sort_values("_ord").drop(columns=["_ord"]),
                 use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════
# TAB 5 — ACTUALIZAÇÕES AUTOMÁTICAS
# ══════════════════════════════════════════════════════
with tab5:
    st.markdown("### 🤖 Actualizações Automáticas de Dados Oficiais")

    col_a, col_b = st.columns([1, 1])
    with col_a:
        st.markdown("""
**Como funciona:**

Um script corre automaticamente **no dia 1 de cada mês** (GitHub Actions, gratuito)
e vai buscar dados a fontes oficiais:

| Fonte | O que faz | Estado |
|-------|-----------|--------|
| 🔴 **DRE** (dre.pt) | Detecta nomeações e exonerações de dirigentes publicadas | ✅ Ativo |
| 🟢 **Transparência Gov** | Actualiza dirigentes superiores da AP | ✅ Ativo |
| 🔵 **dados.gov.pt** | Pesquisa novos datasets de entidades | ✅ Ativo |
| 🟡 **SIOE** | Exportação completa requer login Gov.pt | ⚠️ Manual |
""")

    with col_b:
        if github_configurado():
            st.markdown("**Estado da última execução:**")
            with st.spinner("A verificar log..."):
                log_data = carregar_log_github()

            if log_data:
                ultima = log_data.get("ultima_execucao","?")
                try:
                    dt = datetime.fromisoformat(ultima)
                    ultima_fmt = dt.strftime("%d/%m/%Y às %H:%M")
                except Exception:
                    ultima_fmt = ultima

                st.success(f"✅ Última actualização: **{ultima_fmt}**")

                acts = log_data.get("actualizacoes", [])
                if acts:
                    u = acts[0]
                    st.markdown(f"**Registos:** {u.get('registos_inicial','?')} → {u.get('registos_final','?')}")
                    alertas = u.get("alertas_dre", [])
                    if alertas:
                        st.warning(f"⚠️ **{len(alertas)} possíveis mudanças de dirigentes** detectadas no DRE — confirma e aplica manualmente na Tab '➕ Importar / Actualizar'")
                        with st.expander(f"Ver {len(alertas)} alertas do DRE"):
                            for a in alertas:
                                st.markdown(
                                    f"**{a.get('entidade','')}** — "
                                    f"~~{a.get('dirigente_old','')}~~ → **{a.get('dirigente_new','')}** "
                                    f"(*{a.get('cargo','')}*) "
                                    f"[DRE]({a.get('fonte_dre','')})"
                                )
                    fontes = u.get("fontes", {})
                    if fontes:
                        st.markdown("**Detalhe por fonte:**")
                        for nome_fonte, info in fontes.items():
                            st.markdown(f"- **{nome_fonte}**: {info}")

                    st.markdown("**Histórico recente:**")
                    hist = []
                    for entry in acts[:6]:
                        try:
                            dt2 = datetime.fromisoformat(entry.get("data",""))
                            hist.append({
                                "Data": dt2.strftime("%d/%m/%Y"),
                                "Registos": f"{entry.get('registos_final','?'):,}" if isinstance(entry.get('registos_final'), int) else "?",
                                "Alertas DRE": len(entry.get("alertas_dre",[])),
                            })
                        except Exception:
                            pass
                    if hist:
                        import pandas as _pd
                        st.dataframe(_pd.DataFrame(hist), use_container_width=True, hide_index=True)
            else:
                st.info("Sem log de actualizações ainda. O script corre no dia 1 de cada mês.")
                st.markdown("Podes correr manualmente: **GitHub → Actions → Actualização Automática → Run workflow**")
        else:
            st.warning("GitHub não configurado — actualizações automáticas inativas.")

    st.markdown("---")
    st.markdown("""
**Para activar as actualizações automáticas:**

1. Adiciona nos Secrets do Streamlit Cloud:
```toml
GITHUB_TOKEN = "ghp_o_teu_token"
GITHUB_REPO  = "bernardohille-cmyk/Pesquisar-emails-FP"
```

2. O ficheiro `.github/workflows/auto_update.yml` já está no repositório — não precisas de fazer mais nada.

3. Corre no dia 1 de cada mês automaticamente.

**Limitação importante — SIOE:**
O SIOE exige autenticação Gov.pt para exportar a base completa.
Para actualizar do SIOE tens de fazer o download manual em [sioe.pt](https://www.sioe.dgap.gov.pt)
e importar na Tab "➕ Importar / Actualizar".
""")
