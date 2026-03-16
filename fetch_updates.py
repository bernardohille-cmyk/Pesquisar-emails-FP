"""
AP Contactos — Actualizador Automático de Dados Oficiais
Corre em GitHub Actions (mensalmente).

Fontes:
  1. DRE (dre.pt)          — nomeações e exonerações de dirigentes
  2. dados.gov.pt           — datasets abertos da AP portuguesa
  3. DGAP (dgap.gov.pt)    — base de dirigentes da AP
  4. SIOE (sioe.pt)         — estrutura orgânica (requer download manual, mas aqui tentamos)

Resultado: actualiza dados/contactos.csv e dados/log_atualizacoes.json
"""

import base64
import csv
import io
import json
import os
import re
import sys
import time
import unicodedata
from datetime import datetime, date

import requests

# ── Configuração ──────────────────────────────────────────────
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO  = os.environ.get("GITHUB_REPO", "")
CSV_PATH     = "dados/contactos.csv"
LOG_PATH     = "dados/log_atualizacoes.json"

HEADERS_GH = {
    "Authorization": f"token {GITHUB_TOKEN}",
    "Content-Type":  "application/json",
    "Accept":        "application/vnd.github.v3+json",
}

HEADERS_DRE = {
    "User-Agent":  "AP-Contactos-INA/2.0 (contacto@ina.pt)",
    "Accept":      "application/json",
}


# ── Helpers ───────────────────────────────────────────────────
def limpar(v):
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    return "".join(
        c for c in v if unicodedata.category(c) not in ("Cf", "Cc") or c in ("\n", "\t")
    ).strip()

def email_ok(e):
    e = limpar(str(e)).lower()
    if not e or e in ("nan", "none", "n/d"): return ""
    return e if re.match(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$", e) else ""

def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


# ── GitHub: ler e escrever CSV ─────────────────────────────────
def github_get_file(path):
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"
    r = requests.get(url, headers=HEADERS_GH, timeout=20)
    if r.status_code == 404: return None, None
    r.raise_for_status()
    data = r.json()
    content = base64.b64decode(data["content"]).decode("utf-8")
    return content, data["sha"]

def github_put_file(path, content_str, sha, message):
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"
    payload = {
        "message": message,
        "content": base64.b64encode(content_str.encode("utf-8")).decode(),
    }
    if sha: payload["sha"] = sha
    r = requests.put(url, headers=HEADERS_GH, data=json.dumps(payload), timeout=30)
    r.raise_for_status()
    return r.json().get("content", {}).get("sha")

def carregar_csv_github():
    content, sha = github_get_file(CSV_PATH)
    if content is None:
        return [], sha
    reader = csv.DictReader(io.StringIO(content))
    return list(reader), sha

def guardar_csv_github(rows, sha, mensagem):
    if not rows: return sha
    fieldnames = ["sigla_entidade","designacao","ministerio","tipo_entidade",
                  "orgao_direcao","nome_dirigente","email","contacto",
                  "categoria","fonte","id"]
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=fieldnames, extrasaction="ignore")
    w.writeheader()
    w.writerows(rows)
    novo_sha = github_put_file(CSV_PATH, buf.getvalue(), sha, mensagem)
    log(f"CSV guardado: {len(rows)} registos")
    return novo_sha


# ── 1. DRE — Nomeações e Exonerações ─────────────────────────
def parse_dre_nomeacao(texto):
    """
    Extrai nome e cargo de um texto de despacho de nomeação do DRE.
    Padrões comuns:
      "nomear X para o cargo de Y"
      "designar X como Y"
    """
    nome, cargo = "", ""
    # Padrão: "nomear [Nome] para o cargo de [Cargo]"
    m = re.search(r"nomear\s+([A-ZÀ-Ü][a-zà-ü]+(?:\s+[A-ZÀ-Üa-zà-ü]+){1,5})\s+para\s+(?:o\s+cargo\s+de\s+)?([^,\.]+)", texto, re.IGNORECASE)
    if m:
        nome  = limpar(m.group(1))
        cargo = limpar(m.group(2))
    # Padrão alternativo: "designar [Nome], [Cargo]"
    if not nome:
        m = re.search(r"designar\s+([A-ZÀ-Ü][a-zà-ü]+(?:\s+[A-ZÀ-Üa-zà-ü]+){1,5}),?\s+como\s+([^,\.]+)", texto, re.IGNORECASE)
        if m:
            nome  = limpar(m.group(1))
            cargo = limpar(m.group(2))
    return nome, cargo

def fetch_dre(dias_atras=35):
    """
    Vai ao DRE buscar despachos recentes de nomeação de dirigentes.
    API pública: https://dre.pt/api/v2/
    """
    log("DRE: a pesquisar nomeações recentes...")
    resultados = []
    data_ini = date.today().replace(day=1)  # desde início do mês

    termos = [
        "nomeação dirigente",
        "designar cargo direcção",
        "nomeação presidente",
        "nomeação director-geral",
        "exoneração dirigente",
    ]

    for termo in termos:
        try:
            url = "https://dre.pt/api/v2/search/"
            params = {
                "q":        termo,
                "dateFrom": data_ini.isoformat(),
                "pageSize": 50,
                "type":     "Despacho,Resolução do Conselho de Ministros,Decreto-Lei",
            }
            r = requests.get(url, params=params, headers=HEADERS_DRE, timeout=20)
            if r.status_code != 200:
                log(f"  DRE '{termo}': HTTP {r.status_code}")
                continue
            data = r.json()
            items = data.get("results", [])
            log(f"  DRE '{termo}': {len(items)} documentos")
            for item in items:
                resultados.append({
                    "fonte_origem": "DRE",
                    "titulo":  item.get("title", ""),
                    "data":    item.get("date", ""),
                    "url":     item.get("url", ""),
                    "sumario": item.get("summary", ""),
                    "tipo":    item.get("type", ""),
                })
            time.sleep(0.5)
        except Exception as e:
            log(f"  DRE erro '{termo}': {e}")

    log(f"DRE: {len(resultados)} documentos encontrados")
    return resultados

def processar_dre_para_contactos(docs_dre, contactos_existentes):
    """
    Tenta cruzar documentos DRE com contactos existentes para actualizar dirigentes.
    Retorna (n_actualizados, n_novos_por_confirmar)
    """
    n_upd = 0
    alertas = []

    email_map = {row["email"].lower(): i
                 for i, row in enumerate(contactos_existentes)
                 if row.get("email","").strip()}

    for doc in docs_dre:
        sumario = doc.get("sumario","") + " " + doc.get("titulo","")
        nome, cargo = parse_dre_nomeacao(sumario)
        if not nome: continue

        # Tentar encontrar a entidade mencionada
        for row in contactos_existentes:
            desig = row.get("designacao","").lower()
            # Se o sumário mencionar a designação da entidade
            if len(desig) > 5 and desig in sumario.lower():
                if nome and nome != row.get("nome_dirigente",""):
                    alertas.append({
                        "entidade":      row.get("designacao",""),
                        "dirigente_old": row.get("nome_dirigente",""),
                        "dirigente_new": nome,
                        "cargo":         cargo or row.get("orgao_direcao",""),
                        "fonte_dre":     doc.get("url",""),
                        "data_dre":      doc.get("data",""),
                    })
                    n_upd += 1
                break

    return n_upd, alertas


# ── 2. dados.gov.pt — Datasets abertos ───────────────────────
def fetch_dados_gov():
    """
    Portal de dados abertos português (CKAN).
    Vai buscar datasets relevantes sobre entidades da AP.
    """
    log("dados.gov.pt: a pesquisar datasets...")
    resultados = []
    try:
        # API CKAN
        url = "https://dados.gov.pt/api/3/action/package_search"
        params = {"q": "administração pública entidades dirigentes", "rows": 20}
        r = requests.get(url, params=params, timeout=20)
        if r.status_code != 200:
            log(f"  dados.gov.pt: HTTP {r.status_code}")
            return []
        data = r.json()
        datasets = data.get("result", {}).get("results", [])
        for ds in datasets:
            for resource in ds.get("resources", []):
                fmt = resource.get("format","").upper()
                if fmt in ("CSV","XLSX","XLS","JSON"):
                    resultados.append({
                        "nome":     ds.get("title",""),
                        "url":      resource.get("url",""),
                        "formato":  fmt,
                        "dataset":  ds.get("name",""),
                    })
        log(f"  dados.gov.pt: {len(resultados)} recursos encontrados")
    except Exception as e:
        log(f"  dados.gov.pt erro: {e}")
    return resultados

def fetch_dataset_csv(url, max_rows=5000):
    """Descarrega um CSV de uma URL e devolve lista de dicts."""
    try:
        r = requests.get(url, timeout=30, stream=True)
        r.raise_for_status()
        content = r.content.decode("utf-8", errors="replace")
        reader = csv.DictReader(io.StringIO(content))
        rows = []
        for i, row in enumerate(reader):
            if i >= max_rows: break
            rows.append(dict(row))
        return rows
    except Exception as e:
        log(f"  Erro a descarregar {url}: {e}")
        return []


# ── 3. DGAP — Dirigentes da AP ────────────────────────────────
def fetch_dgap():
    """
    DGAP (Direção-Geral da Administração Pública) — transparencia.gov.pt
    Base pública de dirigentes superiores da AP.
    """
    log("DGAP/transparencia.gov.pt: a pesquisar dirigentes...")
    resultados = []
    try:
        # transparencia.gov.pt tem API de dirigentes
        url = "https://www.transparencia.gov.pt/api/v1/dirigentes"
        params = {"page": 1, "perPage": 500, "ativo": "true"}
        r = requests.get(url, params=params, timeout=20,
                         headers={"User-Agent": "AP-Contactos-INA/2.0"})
        if r.status_code != 200:
            log(f"  transparencia.gov.pt: HTTP {r.status_code}")
            return []
        data = r.json()
        items = data.get("data", data if isinstance(data, list) else [])
        for item in items[:500]:
            resultados.append({
                "nome":       limpar(item.get("nome","") or item.get("name","")),
                "cargo":      limpar(item.get("cargo","") or item.get("position","")),
                "entidade":   limpar(item.get("entidade","") or item.get("entity","")),
                "ministerio": limpar(item.get("ministerio","") or item.get("ministry","")),
                "inicio":     item.get("dataInicio","") or item.get("startDate",""),
                "fonte":      "transparencia.gov.pt",
            })
        log(f"  DGAP: {len(resultados)} dirigentes encontrados")
    except Exception as e:
        log(f"  DGAP erro: {e}")
    return resultados

def cruzar_dgap_com_base(dirigentes_dgap, contactos):
    """
    Cruza dirigentes da DGAP com contactos existentes pelo nome da entidade.
    Actualiza nome_dirigente e orgao_direcao onde encontra correspondência.
    """
    n_upd = 0
    for dig in dirigentes_dgap:
        entidade = dig.get("entidade","").lower().strip()
        if len(entidade) < 5: continue
        for row in contactos:
            desig = row.get("designacao","").lower().strip()
            # Correspondência parcial (nome da entidade contido)
            if (entidade in desig or desig in entidade) and len(desig) > 5:
                nome_novo  = dig.get("nome","").strip()
                cargo_novo = dig.get("cargo","").strip()
                nome_atual = row.get("nome_dirigente","").strip()
                if nome_novo and nome_novo != nome_atual:
                    row["nome_dirigente"]  = nome_novo
                    row["orgao_direcao"]   = cargo_novo or row.get("orgao_direcao","")
                    row["fonte"]           = f"DGAP actualizado {datetime.now().strftime('%Y-%m-%d')}"
                    n_upd += 1
                break
    return n_upd


# ── 4. SIOE — Tentativa de download automático ───────────────
def fetch_sioe_exportacao():
    """
    Tenta descarregar a exportação do SIOE.
    O SIOE exige login (autenticação Gov.pt) para exportar,
    por isso esta função tenta o endpoint público de consulta.
    Se não conseguir, regista no log e continua.
    """
    log("SIOE: a tentar acesso público...")
    try:
        # Endpoint de pesquisa público do SIOE (sem login, dados limitados)
        url = "https://www.sioe.dgap.gov.pt/sioe/pesquisa-entidades"
        params = {"tipoPesquisa": "TODAS", "page": 1, "size": 100}
        r = requests.get(url, params=params, timeout=20,
                         headers={"User-Agent": "AP-Contactos-INA/2.0",
                                  "Accept": "application/json"})
        if r.status_code == 200:
            try:
                data = r.json()
                entidades = data.get("entidades", data.get("content", []))
                log(f"  SIOE público: {len(entidades)} entidades")
                return entidades
            except Exception:
                log("  SIOE: resposta não é JSON (provavelmente HTML — requer login)")
                return []
        else:
            log(f"  SIOE: HTTP {r.status_code} — exportação completa requer login Gov.pt")
            return []
    except Exception as e:
        log(f"  SIOE erro: {e}")
        return []


# ── LOG de actualizações ──────────────────────────────────────
def carregar_log():
    content, sha = github_get_file(LOG_PATH)
    if content is None: return {"actualizacoes": []}, sha
    try: return json.loads(content), sha
    except: return {"actualizacoes": []}, sha

def guardar_log(log_data, sha):
    content = json.dumps(log_data, ensure_ascii=False, indent=2)
    github_put_file(LOG_PATH, content, sha,
                    f"Log actualização {datetime.now().strftime('%Y-%m-%d')}")


# ── MAIN ──────────────────────────────────────────────────────
def main():
    log("=" * 60)
    log(f"AP Contactos — Actualizador automático {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    log("=" * 60)

    if not GITHUB_TOKEN or not GITHUB_REPO:
        log("ERRO: GITHUB_TOKEN e GITHUB_REPO têm de estar definidos.")
        sys.exit(1)

    # Carregar dados actuais
    log("A carregar base de contactos do GitHub...")
    contactos, csv_sha = carregar_csv_github()
    log(f"Base carregada: {len(contactos)} registos")

    resumo = {
        "data": datetime.now().isoformat(),
        "registos_inicial": len(contactos),
        "fontes": {},
        "alertas_dre": [],
    }

    # ── 1. DRE ──
    docs_dre = fetch_dre()
    resumo["fontes"]["DRE"] = {"documentos": len(docs_dre)}
    if docs_dre and contactos:
        n_upd, alertas = processar_dre_para_contactos(docs_dre, contactos)
        resumo["fontes"]["DRE"]["possiveis_atualizacoes"] = n_upd
        resumo["alertas_dre"] = alertas[:20]  # máx 20 alertas
        if alertas:
            log(f"DRE: {len(alertas)} possíveis actualizações de dirigentes detectadas")
            for a in alertas[:5]:
                log(f"  → {a['entidade']}: {a['dirigente_old']} → {a['dirigente_new']}")

    # ── 2. dados.gov.pt ──
    datasets = fetch_dados_gov()
    resumo["fontes"]["dados_gov"] = {"datasets": len(datasets)}
    # Tentar descarregar datasets de entidades/dirigentes
    for ds in datasets[:3]:
        if any(k in ds["nome"].lower() for k in ["dirigente","entidade","organização","administração"]):
            log(f"  dados.gov.pt: a tentar descarregar '{ds['nome']}'...")
            rows = fetch_dataset_csv(ds["url"])
            if rows:
                log(f"  dados.gov.pt: {len(rows)} linhas em '{ds['nome']}'")
            time.sleep(1)

    # ── 3. DGAP/transparencia.gov.pt ──
    dirigentes_dgap = fetch_dgap()
    resumo["fontes"]["DGAP"] = {"dirigentes": len(dirigentes_dgap)}
    if dirigentes_dgap and contactos:
        n_upd_dgap = cruzar_dgap_com_base(dirigentes_dgap, contactos)
        resumo["fontes"]["DGAP"]["actualizacoes"] = n_upd_dgap
        if n_upd_dgap > 0:
            log(f"DGAP: {n_upd_dgap} dirigentes actualizados na base")

    # ── 4. SIOE (tentativa) ──
    entidades_sioe = fetch_sioe_exportacao()
    resumo["fontes"]["SIOE_auto"] = {
        "entidades": len(entidades_sioe),
        "nota": "Exportação completa requer login Gov.pt — fazer manualmente" if not entidades_sioe else "OK"
    }

    # ── Guardar resultados ──
    resumo["registos_final"] = len(contactos)
    resumo["alteracoes"] = len(contactos) - resumo["registos_inicial"]

    if resumo["fontes"].get("DGAP", {}).get("actualizacoes", 0) > 0:
        log("A guardar base actualizada...")
        # Re-numerar IDs se necessário
        for i, row in enumerate(contactos):
            if not row.get("id"): row["id"] = str(i)
        csv_sha = guardar_csv_github(
            contactos, csv_sha,
            f"Actualização automática — DGAP — {datetime.now().strftime('%Y-%m-%d')}"
        )
    else:
        log("Sem alterações à base de contactos nesta execução.")

    # Guardar log
    log_data, log_sha = carregar_log()
    log_data["actualizacoes"].insert(0, resumo)
    log_data["actualizacoes"] = log_data["actualizacoes"][:50]  # manter últimas 50
    log_data["ultima_execucao"] = resumo["data"]
    log_data["fontes_disponiveis"] = {
        "DRE":            "dre.pt/api/v2/ — API pública, nomeações/exonerações",
        "dados_gov":      "dados.gov.pt — CKAN API pública, datasets abertos",
        "DGAP":           "transparencia.gov.pt — dirigentes superiores da AP",
        "SIOE":           "sioe.pt — requer login Gov.pt para exportação completa",
    }
    guardar_log(log_data, log_sha)

    log("=" * 60)
    log(f"Concluído. Registos: {resumo['registos_inicial']} → {resumo['registos_final']}")
    log(f"Alertas DRE (dirigentes possivelmente mudados): {len(resumo['alertas_dre'])}")
    log("=" * 60)

    # Exit code 0 mesmo que não haja alterações
    return 0

if __name__ == "__main__":
    sys.exit(main())
