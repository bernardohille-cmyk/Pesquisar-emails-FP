"""
AP Contactos — Base de Contactos da Administração Pública Portuguesa
Persistência via GitHub. Visual INA.
SIOE importado manualmente (sem API).
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

st.set_page_config(page_title="AP Contactos — INA", page_icon="🏛️", layout="wide")

# ─────────────────────────────────────────────────────
# AUTENTICAÇÃO
# ─────────────────────────────────────────────────────
_PASSWORD = "INA#Contactos2026!"

def check_password():
    if st.session_state.get("autenticado"):
        return
    st.markdown("""
    <div style='max-width:440px;margin:80px auto 0;text-align:center'>
      <div style='font-size:3rem;color:#1F4E79'>🏛️</div>
      <h1 style='color:#1F4E79;margin:8px 0 4px;font-weight:700;letter-spacing:-.5px'>AP Contactos</h1>
      <p style='color:#888;margin:0 0 28px;font-size:.95rem'>
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
# ─────────────────────────────────────────────────────
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
    "Setor Empresarial Regional",
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

# Mapeamento SIOE (coluna "Tipo de Entidade") → Categoria oficial
MAPA_SIOE = {
    "Secretaria-geral": "Administração Direta — Secretarias-Gerais",
    "Direção-geral": "Administração Direta — Direções-Gerais",
    "Inspeção-geral": "Administração Direta — Inspeções-Gerais",
    "Inspeção Regional": "Administração Direta — Inspeções-Gerais",
    "Gabinete 1.º Ministro": "Administração Direta — Gabinetes do Governo",
    "Gabinete Ministro": "Administração Direta — Gabinetes do Governo",
    "Gabinete Secretário de Estado": "Administração Direta — Gabinetes do Governo",
    "Gabinete": "Administração Direta — Gabinetes do Governo",
    "Serviço de Apoio": "Administração Direta — Outros Serviços",
    "Instituto Público": "Administração Indireta — Institutos e Agências",
    "Agência": "Administração Indireta — Institutos e Agências",
    "Centro de Formação Profissional": "Administração Indireta — Institutos e Agências",
    "Banco Central": "Administração Indireta — Institutos e Agências",
    "Fundo Autónomo": "Administração Indireta — Fundos e Fundações",
    "Fundo da Segurança Social": "Administração Indireta — Fundos e Fundações",
    "Entidade Administrativa Independente": "Entidades Reguladoras Independentes",
    "Órgão Independente": "Entidades Reguladoras Independentes",
    "Força de Segurança": "Defesa e Segurança",
    "Forças Armadas": "Defesa e Segurança",
    "Gabinete Secretário Regional": "Regiões Autónomas — Gabinetes",
    "Gabinete Presidente Regional": "Regiões Autónomas — Gabinetes",
    "Gabinete Vice-Presidente Regional": "Regiões Autónomas — Gabinetes",
    "Gabinete do Representante da República": "Regiões Autónomas — Gabinetes",
    "Direção Regional": "Regiões Autónomas — Serviços Regionais",
    "Município": "Poder Local — Municípios",
    "Serviço Municipalizado e Intermunicipalizado": "Poder Local — Municípios",
    "Junta de Freguesia": "Poder Local — Freguesias",
    "Associação de Municípios de fins específicos": "Poder Local — Entidades Intermunicipais",
    "Comunidade intermunicipal": "Poder Local — Entidades Intermunicipais",
    "Federação de Municípios": "Poder Local — Entidades Intermunicipais",
    "Área Metropolitana": "Poder Local — Entidades Intermunicipais",
    "Associação de Freguesias": "Poder Local — Entidades Intermunicipais",
    "Entidade Pública Empresarial": "Setor Empresarial do Estado",
    "Entidade Pública Empresarial Regional": "Setor Empresarial Regional",
    "Entidade Empresarial Regional": "Setor Empresarial Regional",
    "Sociedade Anónima": "Setor Empresarial do Estado",
    "Sociedade por Quotas": "Setor Empresarial do Estado",
    "Cooperativa": "Setor Empresarial do Estado",
    "Agrupamento Complementar de Empresas ": "Setor Empresarial do Estado",
    "Empresa Municipal": "Setor Empresarial Local",
    "Empresa Intermunicipal": "Setor Empresarial Local",
    "Entidade Empresarial Municipal": "Setor Empresarial Local",
    "Tribunal": "Justiça — Tribunais",
    "Unidade Orgânica de Ensino e Investigação": "Ensino Superior e Investigação",
    "Unidade Orgânica de Investigação": "Ensino Superior e Investigação",
    "Estabelecimento de educação e ensino básico e secundário": "Ensino Básico e Secundário",
    "Associação": "Associações",
    "Fundação": "Fundações",
    "Estrutura temporária - comissão": "Estruturas Consultivas e Temporárias",
    "Estrutura temporária - estrutura de missão": "Estruturas Consultivas e Temporárias",
    "Estrutura temporária - grupo de projeto": "Estruturas Consultivas e Temporárias",
    "Estrutura temporária - grupo de trabalho": "Estruturas Consultivas e Temporárias",
    "Órgão consultivo": "Estruturas Consultivas e Temporárias",
    "Entidade Regional de Turismo": "Outros",
    "Estrutura atípica": "Outros",
}

# Mapeamento por nome de sheet (normalizado) → Categoria oficial
# Cobre TODAS as sheets vistas no ficheiro INA (e variantes de acentos/espaços)
MAPA_SHEET = {
    "secretaria-geral governo": "Administração Direta — Secretarias-Gerais",
    "secretarias-gerais": "Administração Direta — Secretarias-Gerais",
    "direcoes-gerais": "Administração Direta — Direções-Gerais",
    "direções-gerais": "Administração Direta — Direções-Gerais",
    "inspecoes-gerais": "Administração Direta — Inspeções-Gerais",
    "inspeções-gerais": "Administração Direta — Inspeções-Gerais",
    "autoridades": "Entidades Reguladoras Independentes",
    "institutos-agencias": "Administração Indireta — Institutos e Agências",
    "institutos-agências": "Administração Indireta — Institutos e Agências",
    "direcoes-regionais": "Regiões Autónomas — Serviços Regionais",
    "direções-regionais": "Regiões Autónomas — Serviços Regionais",
    "empresas publicas": "Setor Empresarial do Estado",
    "empresas públicas": "Setor Empresarial do Estado",
    "sector empresarial regional": "Setor Empresarial Regional",
    "setor empresarial regional": "Setor Empresarial Regional",
    "fundacoes": "Fundações",
    "fundações": "Fundações",
    "comissoes": "Estruturas Consultivas e Temporárias",
    "comissões": "Estruturas Consultivas e Temporárias",
    "academias": "Academias e Associações Científicas",
    "catedras nacionais": "Cátedras UNESCO",
    "cátedras nacionais": "Cátedras UNESCO",
    "camaras municipais": "Poder Local — Municípios",
    "câmaras municipais": "Poder Local — Municípios",
    "comunidades intermunicipais": "Poder Local — Entidades Intermunicipais",
    "associacoes de municipios": "Poder Local — Entidades Intermunicipais",
    "associações de municípios": "Poder Local — Entidades Intermunicipais",
    "servicos municipalizados": "Poder Local — Municípios",
    "serviços municipalizados": "Poder Local — Municípios",
    "juntas de freguesias": "Poder Local — Freguesias",
    "juntas de freguesia": "Poder Local — Freguesias",
    "freguesias": "Poder Local — Freguesias",
    "tribunais": "Justiça — Tribunais",
    "ensino superior": "Ensino Superior e Investigação",
    "universidades": "Ensino Superior e Investigação",
    "politecnicos": "Ensino Superior e Investigação",
    "politécnicos": "Ensino Superior e Investigação",
    "escolas": "Ensino Básico e Secundário",
    "associacoes": "Associações",
    "associações": "Associações",
    "forcas armadas": "Defesa e Segurança",
    "forças armadas": "Defesa e Segurança",
    "forcas de seguranca": "Defesa e Segurança",
    "forças de segurança": "Defesa e Segurança",
}

# Sheets a ignorar (resumos, índices)
SHEETS_IGNORAR = {"menu", "resumo", "conselho editorial", "conselho estratégico",
                  "conselho estrategico", "índice", "indice", "sumário", "sumario"}

COLS_BASE = ["sigla_entidade","designacao","ministerio","tipo_entidade",
             "orgao_direcao","nome_dirigente","email","contacto",
             "categoria","fonte","id"]

GITHUB_DATA_PATH = "dados/contactos.csv"

# ─────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────
def _norm(s):
    if s is None: return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.lower().strip()

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

def categoria_para_ordem(cat):
    try:
        return CATEGORIAS_ORDEM.index(cat)
    except ValueError:
        return 999

def categoria_por_sheet(sheet_name):
    n = _norm(sheet_name)
    if n in MAPA_SHEET:
        return MAPA_SHEET[n]
    # procurar por sub-string
    for k, v in MAPA_SHEET.items():
        if k in n or n in k:
            return v
    return "Outros"

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
    if not repo:
        return None, None
    try:
        url = f"https://api.github.com/repos/{repo}/contents/{GITHUB_DATA_PATH}"
        r = requests.get(url, headers=_github_headers(), timeout=15)
        if r.status_code == 404:
            return None, None
        if r.status_code != 200:
            return None, None
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
    if not repo:
        return False, "GitHub não configurado — adiciona GITHUB_TOKEN e GITHUB_REPO nos Secrets do Streamlit.", sha
    try:
        cols = [c for c in COLS_BASE if c in df.columns]
        csv_str = df[cols].to_csv(index=False)
        content_b64 = base64.b64encode(csv_str.encode("utf-8")).decode()
        payload = {
            "message": f"AP Contactos — actualização {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "content": content_b64,
        }
        if sha:
            payload["sha"] = sha
        url = f"https://api.github.com/repos/{repo}/contents/{GITHUB_DATA_PATH}"
        r = requests.put(url, headers=_github_headers(), data=json.dumps(payload), timeout=30)
        if r.status_code in (200, 201):
            novo_sha = r.json().get("content", {}).get("sha", sha)
            n = len(df)
            n_em = (df["email"].str.len() > 3).sum()
            return True, f"✅ Guardado! {n:,} registos ({n_em:,} emails) na base permanente.", novo_sha
        else:
            return False, f"Erro GitHub {r.status_code}: {r.text[:200]}", sha
    except Exception as e:
        return False, f"Erro ao guardar: {e}", sha

# ─────────────────────────────────────────────────────
# IMPORTADOR TOLERANTE — aceita QUALQUER sheet do Excel
# ─────────────────────────────────────────────────────
# Nomes possíveis (tudo em lowercase, sem acentos) para cada campo
ALIAS_COLUNAS = {
    "designacao": ["designacao", "designação", "nome", "entidade", "organismo",
                   "denominacao", "denominação", "nome da entidade"],
    "sigla_entidade": ["sigla", "acronimo", "acrónimo", "sigla entidade"],
    "ministerio": ["ministerio", "ministério", "tutela", "ministerio/secretaria regional",
                   "ministério/secretaria regional", "secretaria regional"],
    "tipo_entidade": ["tipo", "tipo de entidade", "tipo entidade", "natureza"],
    "orgao_direcao": ["orgao", "órgão", "cargo", "orgao de direcao", "órgão de direção",
                      "funcao", "função", "orgao/cargo"],
    "nome_dirigente": ["nome dirigente", "dirigente", "nome do dirigente", "nome",
                       "responsavel", "responsável", "titular"],
    "email": ["email", "e-mail", "correio electronico", "correio eletrónico",
              "email institucional", "endereco electronico"],
    "contacto": ["contacto", "telefone", "tel", "telf", "nº de telefone",
                 "numero de telefone", "número de telefone"],
}

def _match_coluna(col_name, alvos):
    n = _norm(col_name)
    for a in alvos:
        an = _norm(a)
        if n == an or an in n:
            return True
    return False

def _detectar_colunas(df):
    """Dado um DataFrame com cabeçalhos detectados, devolve mapa campo→coluna real."""
    mapa = {}
    for campo, aliases in ALIAS_COLUNAS.items():
        for col in df.columns:
            if _match_coluna(col, aliases):
                mapa[campo] = col
                break
    return mapa

def _detectar_header(xl, sheet):
    """Lê as primeiras 6 linhas e escolhe a que parece ser o cabeçalho (tem mais campos reconhecíveis)."""
    best_row, best_score = 0, -1
    for h in range(0, 6):
        try:
            tmp = xl.parse(sheet, header=h, nrows=2)
            tmp.columns = [str(c).strip() for c in tmp.columns]
            mapa = _detectar_colunas(tmp)
            score = len(mapa) + (2 if "email" in mapa else 0) + (2 if "designacao" in mapa else 0)
            if score > best_score:
                best_score, best_row = score, h
        except Exception:
            continue
    return best_row, best_score

@st.cache_data(show_spinner=False)
def carregar_base_excel(file_bytes: bytes, filename: str = "") -> pd.DataFrame:
    """Lê todas as sheets do Excel INA de forma tolerante."""
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    frames = []
    sheets_ok, sheets_ignoradas = [], []
    for sheet in xl.sheet_names:
        if _norm(sheet) in SHEETS_IGNORAR:
            sheets_ignoradas.append(sheet)
            continue
        try:
            header_row, score = _detectar_header(xl, sheet)
            if score < 2:
                sheets_ignoradas.append(sheet + " (sem cabeçalho válido)")
                continue
            df = xl.parse(sheet, header=header_row)
            df.columns = [str(c).strip() for c in df.columns]
            mapa = _detectar_colunas(df)
            if "designacao" not in mapa and "nome_dirigente" not in mapa:
                sheets_ignoradas.append(sheet + " (sem coluna de entidade)")
                continue
            novo = pd.DataFrame()
            for campo in ["designacao","sigla_entidade","ministerio","tipo_entidade",
                          "orgao_direcao","nome_dirigente","email","contacto"]:
                col = mapa.get(campo)
                novo[campo] = df[col].apply(limpar) if col and col in df.columns else ""
            novo["email"] = novo["email"].apply(email_ok)
            novo["categoria"] = categoria_por_sheet(sheet)
            novo["fonte"] = f"INA:{sheet}"
            novo = novo[novo["designacao"].str.len() > 2].copy()
            if not novo.empty:
                frames.append(novo)
                sheets_ok.append(f"{sheet} ({len(novo)})")
        except Exception as e:
            sheets_ignoradas.append(f"{sheet} (erro: {e})")
            continue

    if not frames:
        st.session_state["_import_log"] = {"ok": sheets_ok, "ign": sheets_ignoradas}
        return pd.DataFrame(columns=COLS_BASE)

    dfall = pd.concat(frames, ignore_index=True)
    dfall = dfall.reset_index(drop=True)
    dfall["id"] = dfall.index
    st.session_state["_import_log"] = {"ok": sheets_ok, "ign": sheets_ignoradas}
    return garantir_cols(dfall)

# ─────────────────────────────────────────────────────
# IMPORTAR SIOE (manual — ficheiro descarregado do sioe.pt)
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
        email_s = email_s.where(~((t == "email") & (email_s == "")), v.str.lower())
        tel_s   = tel_s.where(~((t == "telefone") & (tel_s == "")), v)
    email_s = email_s.where(
        email_s.str.match(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$"), "")

    nome_s, cargo_s = pd.Series("", index=df.index), pd.Series("", index=df.index)
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
# FUNDIR / ACTUALIZAR
# ─────────────────────────────────────────────────────
def fundir_com_base(df_novo: pd.DataFrame) -> tuple:
    df_novo = garantir_cols(df_novo.copy())
    emails_base = set(st.session_state["df"]["email"].str.lower())
    emails_base.discard("")
    df_novo = df_novo[~df_novo["email"].isin(emails_base) | (df_novo["email"] == "")].copy()
    if df_novo.empty:
        return 0, 0
    id_max = int(st.session_state["df"]["id"].max()) + 1 if len(st.session_state["df"]) else 0
    df_novo = df_novo.reset_index(drop=True)
    df_novo["id"] = df_novo.index + id_max
    n_antes = len(st.session_state["df"])
    st.session_state["df"] = pd.concat(
        [st.session_state["df"], df_novo[COLS_BASE]], ignore_index=True)
    st.session_state["dados_alterados"] = True
    return len(st.session_state["df"]) - n_antes, int((df_novo["email"].str.len() > 3).sum())

def atualizar_com_sioe(df_novo: pd.DataFrame) -> tuple:
    df_novo = garantir_cols(df_novo.copy())
    df_base = st.session_state["df"].copy()
    campos = ["nome_dirigente","orgao_direcao","contacto","ministerio",
              "sigla_entidade","tipo_entidade"]
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
            if v and v not in ("nan","None",""):
                df_base.loc[mb, c] = v
        df_base.loc[mb, "fonte"] = "SIOE (actualizado)"
        n_upd += 1

    df_so = df_novo[df_novo["email"].isin(so_novos) | (df_novo["email"] == "")]
    n_add = 0
    if not df_so.empty:
        id_max = int(df_base["id"].max()) + 1 if len(df_base) else 0
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
    cats = sorted(df["categoria"].dropna().unique(), key=categoria_para_ordem)
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df[cols_ok].to_excel(writer, sheet_name="Todos", index=False)
        for cat in cats:
            sub = df[df["categoria"]==cat][cols_ok]
            if len(sub):
                sub.to_excel(writer, sheet_name=str(cat)[:31], index=False)
    return buf.getvalue()

# ─────────────────────────────────────────────────────
# ESTADO INICIAL
# ─────────────────────────────────────────────────────
defaults = {
    "df": None, "sel": {}, "github_sha": None,
    "dados_alterados": False, "import_msg": None,
    "_file_key": None, "_ext_key": None, "_ext_bytes": None, "_ext_name": None,
    "_gh_tentado": False, "_import_log": None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─────────────────────────────────────────────────────
# CSS — Visual INA
# ─────────────────────────────────────────────────────
st.markdown("""<style>
:root {
  --ina-blue: #1F4E79;
  --ina-blue-dark: #14375B;
  --ina-gold: #C9A961;
  --ina-bg: #F5F7FA;
  --ina-text: #1a1a1a;
  --ina-muted: #6b7280;
}
.stApp { background: var(--ina-bg); }
h1, h2, h3 { color: var(--ina-blue) !important; letter-spacing: -0.3px; font-weight: 700; }
h1 { border-bottom: 3px solid var(--ina-gold); padding-bottom: 8px; }
.stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 2px solid #e5e7eb; }
.stTabs [data-baseweb="tab"] { background: transparent; color: var(--ina-muted); font-weight: 600; padding: 8px 16px; }
.stTabs [aria-selected="true"] { color: var(--ina-blue) !important; border-bottom: 3px solid var(--ina-gold) !important; }
.metric-box {
  background: white; border-radius: 10px; padding: 16px 20px;
  border-left: 4px solid var(--ina-blue); margin-bottom: 10px;
  box-shadow: 0 1px 4px rgba(0,0,0,.06);
}
.metric-box .n { font-size: 2rem; font-weight: 700; color: var(--ina-blue); line-height: 1; }
.metric-box .l { font-size: .82rem; color: var(--ina-muted); margin-top: 4px; }
.email-box {
  background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 8px;
  padding: 12px 16px; font-family: 'Consolas', monospace; font-size: .84rem;
  max-height: 300px; overflow-y: auto; white-space: pre-wrap; word-break: break-all;
}
.stButton > button[kind="primary"] {
  background: var(--ina-blue); border: none; font-weight: 600;
}
.stButton > button[kind="primary"]:hover { background: var(--ina-blue-dark); }
.stSidebar { background: linear-gradient(180deg, #FFFFFF 0%, #F0F4F8 100%); }
.ina-header {
  background: linear-gradient(135deg, var(--ina-blue) 0%, var(--ina-blue-dark) 100%);
  color: white; padding: 20px 24px; border-radius: 10px; margin-bottom: 18px;
  border-bottom: 4px solid var(--ina-gold);
  display: flex; align-items: center; gap: 16px;
}
.ina-header .brasao { font-size: 2.2rem; }
.ina-header h1 { color: white !important; margin: 0; border: none; padding: 0; font-size: 1.6rem; }
.ina-header .sub { color: #cbd5e1; font-size: .88rem; margin-top: 2px; }
</style>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏛️ AP Contactos")
    st.caption("Instituto Nacional de Administração")
    st.markdown("---")

    # Estado da ligação ao GitHub
    if github_configurado():
        st.success("🔗 Persistência GitHub activa")
        if st.session_state["df"] is None and not st.session_state["_gh_tentado"]:
            st.session_state["_gh_tentado"] = True
            with st.spinner("A carregar base permanente..."):
                df_gh, sha = carregar_do_github()
                if df_gh is not None:
                    st.session_state["df"] = df_gh
                    st.session_state["github_sha"] = sha
                    st.session_state["dados_alterados"] = False
    else:
        st.warning("⚠️ Persistência desactivada\n\nAdiciona nos Secrets do Streamlit:\nGITHUB_TOKEN e GITHUB_REPO")

    st.markdown("---")
    st.markdown("### 📂 Carregar base Excel (INA)")
    st.caption("Carrega pela primeira vez ou substitui os dados actuais.")
    uploaded = st.file_uploader("Base INA (.xlsx)", type=["xlsx"], label_visibility="collapsed")
    if uploaded is not None:
        file_bytes = uploaded.getvalue()
        file_key = f"{uploaded.name}_{len(file_bytes)}"
        if st.session_state["_file_key"] != file_key:
            st.session_state["_file_key"] = file_key
            st.session_state["sel"] = {}
            st.session_state["import_msg"] = None
            st.session_state["_ext_key"] = None
            st.cache_data.clear()
            with st.spinner("A processar ficheiro Excel..."):
                st.session_state["df"] = carregar_base_excel(file_bytes, uploaded.name)
                st.session_state["dados_alterados"] = True
                st.rerun()

    if st.session_state["df"] is not None:
        n_tot = len(st.session_state["df"])
        n_mail = (st.session_state["df"]["email"].str.len() > 3).sum()
        st.success(f"**{n_tot:,}** registos · **{n_mail:,}** emails")

        if st.session_state["dados_alterados"] and github_configurado():
            if st.button("💾 Guardar na base permanente", type="primary", use_container_width=True):
                with st.spinner("A guardar no GitHub..."):
                    ok, msg, novo_sha = guardar_no_github(
                        st.session_state["df"], st.session_state["github_sha"])
                if ok:
                    st.session_state["github_sha"] = novo_sha
                    st.session_state["dados_alterados"] = False
                    st.session_state["import_msg"] = {"texto": msg}
                    st.rerun()
                else:
                    st.error(msg)

    # log do último import
    if st.session_state.get("_import_log"):
        with st.expander("🔍 Log da última importação"):
            log = st.session_state["_import_log"]
            st.markdown(f"**Sheets lidas ({len(log['ok'])}):**")
            for s in log["ok"]:
                st.caption(f"✓ {s}")
            if log["ign"]:
                st.markdown(f"**Ignoradas ({len(log['ign'])}):**")
                for s in log["ign"]:
                    st.caption(f"× {s}")

    st.markdown("---")
    st.markdown("### 🔍 Filtros")

# ─────────────────────────────────────────────────────
# GUARD
# ─────────────────────────────────────────────────────
if st.session_state["df"] is None:
    st.markdown("""
    <div class='ina-header'>
      <div class='brasao'>🏛️</div>
      <div>
        <h1>Contactos — AP Portuguesa</h1>
        <div class='sub'>Instituto Nacional de Administração</div>
      </div>
    </div>
    """, unsafe_allow_html=True)
    if github_configurado():
        st.info("A base permanente parece estar vazia. Carrega o ficheiro Excel do INA na barra lateral para começar.")
    else:
        st.info("👆 Carrega o ficheiro Excel do INA na barra lateral para começar.")
        st.markdown("""
        **Para activar a persistência automática** (os dados deixam de desaparecer), adiciona nos *Secrets* do Streamlit Cloud:

        ```toml
        GITHUB_TOKEN = "ghp_o_teu_personal_access_token"
        GITHUB_REPO  = "bernardohille-cmyk/Pesquisar-emails-FP"
        ```

        O token deve ter permissão **Contents: Read and write** no repositório.
        """)
    st.stop()

df = st.session_state["df"]

# ─────────────────────────────────────────────────────
# FILTROS SIDEBAR
# ─────────────────────────────────────────────────────
with st.sidebar:
    pesquisa = st.text_input("🔎 Pesquisar", "")
    cats_disp = sorted(df["categoria"].dropna().unique(), key=categoria_para_ordem)
    cats_sel  = st.multiselect("Categoria", cats_disp)
    _prefixos = ("Ministério","Secretaria","Presidência","Vice-Presidência")
    min_disp = sorted(m for m in df["ministerio"].fillna("").unique()
                      if m and any(m.startswith(p) for p in _prefixos))
    min_sel  = st.multiselect("Ministério / Tutela", min_disp)
    so_email = st.checkbox("Só com email", value=False)

    st.markdown("---")
    if st.button("🗑️ Limpar sessão & recarregar", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.cache_data.clear()
        st.rerun()

df_f = df.copy()
if pesquisa:
    p = pesquisa.lower()
    m = (df_f["designacao"].str.lower().str.contains(p, na=False) |
         df_f["sigla_entidade"].str.lower().str.contains(p, na=False) |
         df_f["ministerio"].str.lower().str.contains(p, na=False) |
         df_f["nome_dirigente"].str.lower().str.contains(p, na=False) |
         df_f["email"].str.lower().str.contains(p, na=False))
    df_f = df_f[m]
if cats_sel: df_f = df_f[df_f["categoria"].isin(cats_sel)]
if min_sel:  df_f = df_f[df_f["ministerio"].isin(min_sel)]
if so_email: df_f = df_f[df_f["email"].str.len() > 3]

tem_filtro = bool(pesquisa or cats_sel or min_sel or so_email)

# ─────────────────────────────────────────────────────
# CABEÇALHO
# ─────────────────────────────────────────────────────
st.markdown("""
<div class='ina-header'>
  <div class='brasao'>🏛️</div>
  <div>
    <h1>Contactos — AP Portuguesa</h1>
    <div class='sub'>Instituto Nacional de Administração</div>
  </div>
</div>
""", unsafe_allow_html=True)

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
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Tabela", "☑️ Selecionar & Copiar Emails",
    "➕ Importar / Actualizar", "📥 Descarregar",
])
# ══════════════════════════════════════════════════════
# TAB 1 — TABELA
# ══════════════════════════════════════════════════════
with tab1:
    lbl_map = {"categoria":"Categoria","sigla_entidade":"Sigla","designacao":"Entidade",
               "ministerio":"Ministério","orgao_direcao":"Órgão/Cargo",
               "nome_dirigente":"Dirigente","email":"Email","contacto":"Contacto","fonte":"Fonte"}
    cols_v = [c for c in lbl_map if c in df_f.columns]
    st.dataframe(df_f[cols_v].rename(columns=lbl_map),
                 use_container_width=True, height=520, hide_index=True)
    st.caption(f"A mostrar {len(df_f):,} de {len(df):,} registos totais")

    df_em_f = df_f[df_f["email"].str.len() > 3].drop_duplicates(subset=["email"])
    n_em_f = len(df_em_f)
    if n_em_f > 0:
        st.markdown("---")
        st.markdown(f"**📧 Exportação rápida — {n_em_f:,} emails:**")
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
                for i in df_ce["id"].tolist():
                    st.session_state["sel"][i] = True
                st.rerun()
        with c2b:
            if st.button("❌ Limpar", use_container_width=True):
                st.session_state["sel"] = {}
                st.rerun()
        with c3b:
            cat_r = st.selectbox("Adicionar categoria:",
                ["—"] + sorted(df_ce["categoria"].unique(), key=categoria_para_ordem),
                key="cat_r")
            if cat_r != "—" and st.button(f"➕ '{cat_r}'", use_container_width=True):
                for i in df_ce[df_ce["categoria"]==cat_r]["id"].tolist():
                    st.session_state["sel"][i] = True
                st.rerun()

        st.markdown("---")
        for cat in sorted(df_ce["categoria"].unique(), key=categoria_para_ordem):
            sub = df_ce[df_ce["categoria"]==cat]
            n_s = sum(1 for i in sub["id"] if st.session_state["sel"].get(i, False))
            with st.expander(f"**{cat}** — {len(sub)} com email | ✓ {n_s} seleccionados"):
                sc1, sc2 = st.columns(2)
                with sc1:
                    if st.button("✅ Todos", key=f"sa_{cat}", use_container_width=True):
                        for i in sub["id"].tolist(): st.session_state["sel"][i] = True
                        st.rerun()
                with sc2:
                    if st.button("❌ Nenhum", key=f"sn_{cat}", use_container_width=True):
                        for i in sub["id"].tolist(): st.session_state["sel"][i] = False
                        st.rerun()
                for _, row in sub.iterrows():
                    idx = row["id"]
                    lbl_chk = f"{row['designacao']} — {row['email']}"
                    if str(row.get("nome_dirigente","")).strip():
                        lbl_chk += f" ({row['nome_dirigente']})"
                    v = st.checkbox(lbl_chk, value=st.session_state["sel"].get(idx,False),
                                    key=f"chk_{idx}")
                    if v != st.session_state["sel"].get(idx, False):
                        st.session_state["sel"][idx] = v

        st.markdown("---")
        ids_sel = [i for i,v in st.session_state["sel"].items() if v]
        df_sel = df_ce[df_ce["id"].isin(ids_sel)]
        st.markdown(f"### 📋 {len(df_sel)} seleccionados")
        if not df_sel.empty:
            cf1, cf2 = st.columns([2,1])
            with cf1:
                fmt = st.radio("Formato", ["Um por linha","Separados por ; (BCC)","CSV"], horizontal=True)
            with cf2:
                inc_nome = st.checkbox("Incluir nome", value=False)
                dedup = st.checkbox("Sem duplicados", value=True)
            df_exp = df_sel.drop_duplicates(subset=["email"]) if dedup else df_sel
            lista = sorted(df_exp["email"].tolist())
            if inc_nome:
                pares = df_exp[["nome_dirigente","email"]]
                if fmt == "Um por linha":
                    texto = "\n".join(f"{r['nome_dirigente']} <{r['email']}>" for _,r in pares.iterrows())
                elif fmt == "Separados por ; (BCC)":
                    texto = "; ".join(f"{r['nome_dirigente']} <{r['email']}>" for _,r in pares.iterrows())
                else:
                    texto = "Nome,Email\n" + "\n".join('"' + str(r['nome_dirigente']) + '",' + str(r['email']) for _,r in pares.iterrows())
            else:
                if fmt == "Um por linha": texto = "\n".join(lista)
                elif fmt == "Separados por ; (BCC)": texto = "; ".join(lista)
                else: texto = "Email\n" + "\n".join(lista)
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
    st.caption("Suporta exportação SIOE (auto), ou qualquer Excel com mapeamento manual de colunas.")

    uploaded_ext = st.file_uploader("Carregar ficheiro externo (.xlsx)",
                                    type=["xlsx"], key="uploader_ext",
                                    label_visibility="collapsed")
    if uploaded_ext is not None:
        ext_key = f"{uploaded_ext.name}_{uploaded_ext.size}"
        if st.session_state["_ext_key"] != ext_key:
            st.session_state["_ext_key"] = ext_key
            st.session_state["_ext_bytes"] = uploaded_ext.getvalue()
            st.session_state["_ext_name"]  = uploaded_ext.name

    ext_bytes = st.session_state["_ext_bytes"]
    ext_name  = st.session_state["_ext_name"] or ""

    if ext_bytes is None:
        st.info("**Como usar:**\n"
                "- 🏛️ **Exportação SIOE** (ExportResultadosPesquisa*.xlsx): detecção automática + categorias oficiais.\n"
                "- 📋 **Qualquer outro Excel**: mapeamento manual de colunas.")
    else:
        parece_sioe = "export" in ext_name.lower() and "pesquisa" in ext_name.lower()
        if not parece_sioe:
            try:
                _t = pd.read_excel(io.BytesIO(ext_bytes), header=0, nrows=1)
                _t.columns = [str(c).strip() for c in _t.columns]
                parece_sioe = "Designação" in _t.columns and "Código SIOE" in _t.columns
            except Exception:
                pass

        st.markdown(f"**Ficheiro:** `{ext_name}`")
        if st.button("🗑️ Remover", key="rm_ext"):
            st.session_state["_ext_key"] = None
            st.session_state["_ext_bytes"] = None
            st.session_state["_ext_name"] = None
            st.rerun()
        st.markdown("---")

        # ── SIOE ──
        if parece_sioe:
            st.success("✅ Ficheiro SIOE detectado — categorias oficiais atribuídas automaticamente")
            try:
                n_lin = len(pd.read_excel(io.BytesIO(ext_bytes), header=0, usecols=[0]))
            except Exception:
                n_lin = "?"
            cm, ci = st.columns([1,2])
            with cm:
                st.metric("Entidades", f"{n_lin:,}" if isinstance(n_lin,int) else n_lin)
                modo_atualizacao = st.toggle("🔄 Modo actualização", value=False,
                    help="ON: sobrescreve dirigente/cargo/telefone dos registos existentes. OFF: só acrescenta emails novos.")
                btn_sioe = st.button("🚀 Importar SIOE", type="primary", use_container_width=True)
            with ci:
                if modo_atualizacao:
                    st.warning("**Modo actualização** — sobrescreve campos dos registos existentes.", icon="🔄")
                else:
                    st.info("**Modo normal** — só acrescenta novos emails.")

            if btn_sioe:
                with st.spinner("A processar SIOE…"):
                    try:
                        df_novo = importar_sioe(ext_bytes)
                        if df_novo.empty:
                            st.error("Nenhum registo válido.")
                        elif modo_atualizacao:
                            n_upd, n_add = atualizar_com_sioe(df_novo)
                            st.session_state["import_msg"] = {"texto":
                                f"🔄 **{n_upd:,}** registos actualizados, **{n_add:,}** novos. "
                                f"Base: **{len(st.session_state['df']):,}** registos. "
                                f"{'💾 Não te esqueças de guardar!' if github_configurado() else ''}"}
                            st.rerun()
                        else:
                            n_add, n_em = fundir_com_base(df_novo)
                            if n_add == 0:
                                st.warning("Todos os emails do SIOE já existem na base.")
                            else:
                                st.session_state["import_msg"] = {"texto":
                                    f"✅ **{n_add:,}** registos adicionados (**{n_em:,}** com email). "
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
                with cs1:
                    sheet_esc = st.selectbox("Sheet", xl_ext.sheet_names, key="m_sheet")
                with cs2:
                    hrow = int(st.number_input("Linha cabeçalho (0=primeira)", 0, 10, 0, key="m_hrow"))
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
                        if esc != "— (não mapear)":
                            mapa[campo] = esc
                cat_nome = st.selectbox("Categoria oficial", CATEGORIAS_ORDEM, key="m_cat")

                if st.button("🔀 Fundir com a base", type="primary", use_container_width=True, key="btn_manual"):
                    if "designacao" not in mapa or "email" not in mapa:
                        st.error("Tens de mapear pelo menos Nome e Email.")
                    else:
                        try:
                            df_m = xl_ext.parse(sheet_esc, header=hrow)
                            df_m.columns = [str(c).strip() for c in df_m.columns]
                            novo = pd.DataFrame()
                            for campo in ["designacao","sigla_entidade","ministerio","tipo_entidade",
                                          "orgao_direcao","nome_dirigente","email","contacto"]:
                                col_e = mapa.get(campo,"")
                                novo[campo] = df_m[col_e].apply(limpar) if col_e and col_e in df_m.columns else ""
                            novo["email"] = novo["email"].apply(email_ok)
                            novo["categoria"] = cat_nome
                            novo["fonte"] = f"importado:{ext_name}"
                            novo = novo[novo["designacao"].str.len() > 2].copy()
                            if novo.empty:
                                st.warning("Nenhum registo válido.")
                            else:
                                n_add, n_em = fundir_com_base(novo)
                                if n_add == 0:
                                    st.warning("Todos os registos já existem na base.")
                                else:
                                    st.session_state["import_msg"] = {"texto":
                                        f"✅ **{n_add:,}** registos adicionados em **{cat_nome}** "
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
    st.info(f"Base actual: **{n_tot:,}** registos | **{n_mail:,}** com email | **{n_tot-n_mail:,}** sem email")

    cd1, cd2, cd3 = st.columns(3)
    with cd1:
        st.markdown("**Excel completo** (sheet por categoria)")
        st.download_button("⬇️ Excel completo",
            data=df_to_excel_bytes(df_dl),
            file_name=f"AP_Contactos_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary")
    with cd2:
        st.markdown(f"**Excel filtrado** ({len(df_f):,} registos)")
        if not tem_filtro:
            st.caption("Activa um filtro primeiro.")
        else:
            st.download_button(f"⬇️ Excel filtrado ({len(df_f):,})",
                data=df_to_excel_bytes(df_f.drop(columns=["id"],errors="ignore")),
                file_name=f"AP_Contactos_filtrado_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary")
    with cd3:
        st.markdown(f"**CSV filtrado** ({len(df_f):,} registos)")
        if not tem_filtro:
            st.caption("Activa um filtro primeiro.")
        else:
            csv_b = (df_f.drop(columns=["id"],errors="ignore")
                     .to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"))
            st.download_button(f"⬇️ CSV filtrado ({len(df_f):,})",
                data=csv_b,
                file_name=f"AP_Contactos_filtrado_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv", use_container_width=True)

    st.markdown("---")
    st.markdown("**Cobertura por categoria:**")
    res = (df_dl.groupby("categoria")
           .agg(Total=("id","count"),
                Com_Email=("email", lambda x:(x.str.len()>3).sum()))
           .reset_index())
    res.columns = ["Categoria","Total","Com Email"]
    res["Sem Email"] = res["Total"] - res["Com Email"]
    res["Cobertura"] = (res["Com Email"]/res["Total"]*100).round(1).astype(str)+"%"
    res["_ord"] = res["Categoria"].apply(categoria_para_ordem)
    st.dataframe(res.sort_values("_ord").drop(columns=["_ord"]),
                 use_container_width=True, hide_index=True)
