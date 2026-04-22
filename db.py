"""
db.py — Modelo de dados AP Contactos (INA)
2 tabelas: entidades + dirigentes. Persistência via GitHub API.
"""
from __future__ import annotations

import base64
import io
import json
from pathlib import Path
import re
import unicodedata
from datetime import datetime
from typing import Optional

import pandas as pd
import requests
import streamlit as st

# ═══════════════════════════════════════════════════
# CAMINHOS GITHUB
# ═══════════════════════════════════════════════════
PATH_ENTIDADES  = "dados/entidades.csv"
PATH_DIRIGENTES = "dados/dirigentes.csv"
PATH_CATEGORIAS = "dados/categorias.json"
PATH_PENDENTES  = "dados/pendentes.json"
BASE_DIR = Path(__file__).resolve().parent

COLS_ENTIDADE = [
    "entity_id", "designacao", "siglas", "sioe_code", "nif",
    "categoria_id", "ministerio", "tutela",
    "website", "morada", "distrito",
    "estado", "criado_em", "alterado_em", "notas",
]
COLS_DIRIGENTE = [
    "dirig_id", "entity_id", "nome", "cargo",
    "email", "telefone",
    "inicio", "fim",
    "fonte", "confianca",
    "criado_em", "alterado_em", "notas",
]

# ═══════════════════════════════════════════════════
# NORMALIZAÇÃO / SLUGS
# ═══════════════════════════════════════════════════
def norm(s) -> str:
    if s is None: return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.lower().strip()

def slugify(s: str, max_len: int = 60) -> str:
    s = norm(s)
    s = re.sub(r"[^a-z0-9]+", "-", s)
    s = s.strip("-")
    return s[:max_len] or "entidade"

def limpar_str(v) -> str:
    if v is None: return ""
    if isinstance(v, float) and pd.isna(v): return ""
    s = str(v)
    s = "".join(c for c in s if unicodedata.category(c) not in ("Cf","Cc") or c in ("\n","\t"))
    return s.strip()

def email_ok(e) -> str:
    e = limpar_str(e).lower()
    if not e or e in ("nan","none","n/d","-","—"): return ""
    return e if re.match(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$", e) else ""

def tel_ok(t) -> str:
    t = limpar_str(t)
    if not t or t.lower() in ("nan","none","n/d","-","—"): return ""
    return re.sub(r"[^\d+\s()\-]", "", t).strip()

# ═══════════════════════════════════════════════════
# GITHUB I/O
# ═══════════════════════════════════════════════════
def _secret(name: str, default: str = "") -> str:
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default

def gh_cfg():
    return bool(_secret("GITHUB_TOKEN")) and bool(_secret("GITHUB_REPO"))

def _gh_headers():
    token = _secret("GITHUB_TOKEN")
    return {"Authorization": "token " + token, "Accept": "application/vnd.github+json"}

def _gh_repo():
    return _secret("GITHUB_REPO")

def _gh_contents_url(path: str) -> str:
    return "https://api.github.com/repos/" + _gh_repo() + "/contents/" + path

def _local_path(path: str) -> Path:
    return BASE_DIR / path

def _local_get(path: str):
    try:
        p = _local_path(path)
        if not p.exists():
            return None, None
        return p.read_bytes(), None
    except Exception:
        return None, None

def _local_put(path: str, content: bytes):
    try:
        p = _local_path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(content)
        return True, "OK", None
    except Exception as e:
        return False, "Excepção local: " + str(e), None

def gh_get(path):
    """Lê um ficheiro do GitHub ou, sem secrets, do disco local."""
    if not gh_cfg():
        return _local_get(path)
    try:
        url = _gh_contents_url(path)
        r = requests.get(url, headers=_gh_headers(), timeout=15)
        if r.status_code != 200:
            return None, None
        data = r.json()
        sha = data.get("sha")
        content = data.get("content")
        encoding = data.get("encoding")
        if content and encoding == "base64":
            return base64.b64decode(content), sha

        # Ficheiros >1 MB vêm sem conteúdo inline; pedir o raw diretamente.
        raw_headers = _gh_headers() | {"Accept": "application/vnd.github.raw+json"}
        raw_resp = requests.get(url, headers=raw_headers, timeout=30)
        if raw_resp.status_code == 200 and raw_resp.content:
            return raw_resp.content, sha

        download_url = data.get("download_url")
        if download_url:
            raw_resp = requests.get(download_url, timeout=30)
            if raw_resp.status_code == 200:
                return raw_resp.content, sha
        return None, sha
    except Exception:
        return None, None

def gh_put(path, content, sha, mensagem):
    """Escreve um ficheiro no GitHub ou, sem secrets, no disco local."""
    if not gh_cfg():
        return _local_put(path, content)
    try:
        url = _gh_contents_url(path)
        def _put(current_sha):
            payload = {"message": mensagem, "content": base64.b64encode(content).decode()}
            if current_sha:
                payload["sha"] = current_sha
            return requests.put(url, headers=_gh_headers(), data=json.dumps(payload), timeout=30)

        r = _put(sha)
        if r.status_code in (200, 201):
            novo_sha = r.json().get("content", {}).get("sha", sha)
            return True, "OK", novo_sha

        # Sha antigo na sessão: recarregar o ficheiro remoto e tentar uma vez mais.
        if r.status_code == 409:
            latest_raw, latest_sha = gh_get(path)
            if latest_sha:
                if latest_raw == content:
                    return True, "OK", latest_sha
                r = _put(latest_sha)
                if r.status_code in (200, 201):
                    novo_sha = r.json().get("content", {}).get("sha", latest_sha)
                    return True, "OK", novo_sha
        return False, "Erro GitHub " + str(r.status_code) + ": " + r.text[:200], sha
    except Exception as e:
        return False, "Excepção: " + str(e), sha

# ═══════════════════════════════════════════════════
# LOAD / SAVE TABELAS
# ═══════════════════════════════════════════════════
def _df_vazio(cols):
    return pd.DataFrame({c: pd.Series(dtype="object") for c in cols})

def load_entidades():
    raw, sha = gh_get(PATH_ENTIDADES)
    if raw is None:
        return _df_vazio(COLS_ENTIDADE), None
    try:
        df = pd.read_csv(io.BytesIO(raw), dtype=str).fillna("")
        for c in COLS_ENTIDADE:
            if c not in df.columns: df[c] = ""
        return df[COLS_ENTIDADE], sha
    except Exception:
        return _df_vazio(COLS_ENTIDADE), sha

def load_dirigentes():
    raw, sha = gh_get(PATH_DIRIGENTES)
    if raw is None:
        return _df_vazio(COLS_DIRIGENTE), None
    try:
        df = pd.read_csv(io.BytesIO(raw), dtype=str).fillna("")
        for c in COLS_DIRIGENTE:
            if c not in df.columns: df[c] = ""
        return df[COLS_DIRIGENTE], sha
    except Exception:
        return _df_vazio(COLS_DIRIGENTE), sha

def save_entidades(df, sha):
    csv_str = df[COLS_ENTIDADE].to_csv(index=False)
    return gh_put(PATH_ENTIDADES, csv_str.encode("utf-8"), sha,
                  "Entidades: " + str(len(df)) + " registos (" + datetime.now().strftime('%Y-%m-%d %H:%M') + ")")

def save_dirigentes(df, sha):
    csv_str = df[COLS_DIRIGENTE].to_csv(index=False)
    return gh_put(PATH_DIRIGENTES, csv_str.encode("utf-8"), sha,
                  "Dirigentes: " + str(len(df)) + " registos (" + datetime.now().strftime('%Y-%m-%d %H:%M') + ")")

# ═══════════════════════════════════════════════════
# CATEGORIAS
# ═══════════════════════════════════════════════════
_CAT_CACHE = {"data": None}

def load_categorias():
    if _CAT_CACHE["data"] is not None: return _CAT_CACHE["data"]
    raw, _ = gh_get(PATH_CATEGORIAS)
    if raw is None:
        data = {"versao":"0","categorias":[
            {"id":"outros","nome":"Outros","ordem":99,"cor":"#595959","sheets":[],"tipos_sioe":[]}]}
    else:
        try: data = json.loads(raw.decode("utf-8"))
        except Exception: data = {"versao":"0","categorias":[]}
    _CAT_CACHE["data"] = data
    return data

def cat_ordem_map():
    return {c["id"]: c["ordem"] for c in load_categorias()["categorias"]}

def cat_nome(cat_id):
    for c in load_categorias()["categorias"]:
        if c["id"] == cat_id: return c["nome"]
    return cat_id or "—"

def cat_cor(cat_id):
    for c in load_categorias()["categorias"]:
        if c["id"] == cat_id: return c.get("cor","#595959")
    return "#595959"

def cat_por_sheet(sheet_name):
    n = norm(sheet_name)
    for c in load_categorias()["categorias"]:
        for s in c.get("sheets", []):
            if norm(s) == n:
                return c["id"]
    for c in load_categorias()["categorias"]:
        for s in c.get("sheets", []):
            sn = norm(s)
            if sn and (sn == n or (len(sn) > 5 and sn in n)):
                return c["id"]
    return "outros"

def cat_por_tipo_sioe(tipo):
    n = norm(tipo)
    for c in load_categorias()["categorias"]:
        for t in c.get("tipos_sioe", []):
            if norm(t) == n:
                return c["id"]
    return "outros"

# ═══════════════════════════════════════════════════
# CHAVE DE DEDUPLICAÇÃO
# ═══════════════════════════════════════════════════
def chave_entidade(designacao, sigla="", sioe="", nif=""):
    if sioe and str(sioe).strip() and str(sioe).strip().lower() != "nan":
        return "sioe:" + str(sioe).strip()
    if nif and str(nif).strip() and str(nif).strip().lower() != "nan":
        return "nif:" + str(nif).strip()
    d = norm(designacao)
    s = norm(sigla)
    d = re.sub(r"\b(i\.? ?p\.?|e\.? ?p\.? ?e\.?|s\.? ?a\.?|lda\.?|cripp)\b", "", d).strip()
    if s and len(s) >= 2 and s != d:
        return "desig:" + d + "|sigla:" + s
    return "desig:" + d

def chave_dirigente(entity_id, cargo, email="", nome=""):
    if email:
        return entity_id + "|email:" + norm(email)
    return entity_id + "|cargo:" + norm(cargo) + "|nome:" + norm(nome)
