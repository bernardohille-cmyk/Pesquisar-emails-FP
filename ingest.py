"""
ingest.py — Importadores de dados para AP Contactos (INA)

Fontes suportadas:
- Excel INA (com várias sheets por categoria)
- SIOE (manual, ficheiro ExportResultadosPesquisa)
- DRE (scraping de dre.pt para nomeações/exonerações) — devolve SUGESTÕES

Todas as funções devolvem DataFrames com as colunas de entidades/dirigentes,
prontos para o merge controlado em db.py.
"""
from __future__ import annotations

import io
import re
import time
import uuid
from datetime import datetime
from typing import Optional

import pandas as pd
import requests

from db import (
    COLS_ENTIDADE, COLS_DIRIGENTE, load_categorias,
    cat_por_sheet, cat_por_tipo_sioe,
    chave_entidade, chave_dirigente,
    norm, slugify, limpar_str, email_ok, tel_ok,
)

# ═══════════════════════════════════════════════════
# AUTO-DETECÇÃO DE COLUNAS
# ═══════════════════════════════════════════════════
ALIAS = {
    "designacao":    ["designacao","designação","nome","entidade","organismo","denominacao","denominação","nome da entidade"],
    "sigla":         ["sigla","acronimo","acrónimo"],
    "sioe_code":     ["codigo sioe","código sioe","sioe","cod sioe","cód. sioe"],
    "nif":           ["nif","npc"],
    "ministerio":    ["ministerio","ministério","tutela","ministério/secretaria regional","secretaria regional"],
    "tipo_entidade": ["tipo","tipo de entidade","tipo entidade","natureza"],
    "cargo":         ["cargo","orgao","órgão","orgao de direcao","órgão de direção","funcao","função","orgao/cargo","orgao_direcao"],
    "nome_dirigente":["nome dirigente","dirigente","nome do dirigente","responsavel","responsável","titular","membro nome","nome"],
    "email":         ["email","e-mail","correio electronico","correio eletrónico","email institucional"],
    "telefone":      ["contacto","telefone","tel","telf","nº de telefone"],
    "website":       ["website","url","site","sitio","sítio"],
    "morada":        ["morada","endereco","endereço"],
    "distrito":      ["distrito","concelho"],
}

def _match(col_name, alvos):
    n = norm(col_name)
    for a in alvos:
        an = norm(a)
        if n == an or (len(an) >= 4 and an in n):
            return True
    return False

def _detectar_colunas(df):
    m = {}
    for campo, alvos in ALIAS.items():
        for col in df.columns:
            if _match(col, alvos):
                m[campo] = col
                break
    return m

def _detectar_header(xl, sheet):
    best_row, best_score = 0, -1
    for h in range(0, 6):
        try:
            tmp = xl.parse(sheet, header=h, nrows=2)
            tmp.columns = [str(c).strip() for c in tmp.columns]
            m = _detectar_colunas(tmp)
            score = len(m) + (3 if "designacao" in m else 0) + (2 if "email" in m else 0)
            if score > best_score:
                best_score, best_row = score, h
        except Exception:
            continue
    return best_row, best_score

# ═══════════════════════════════════════════════════
# IMPORT EXCEL INA
# ═══════════════════════════════════════════════════
SHEETS_IGNORAR = {"menu","resumo","conselho editorial","conselho estrategico",
                  "conselho estratégico","indice","índice","sumario","sumário",
                  "total","totais","geral"}

def importar_excel_ina(file_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """
    Lê um Excel do INA (multi-sheet por categoria).
    Devolve (df_entidades, df_dirigentes, log).
    
    Regras:
    - Ignora sheets de resumo/índice.
    - Auto-detecta cabeçalho e colunas.
    - Cada linha gera 1 entidade + 1 dirigente (se houver).
    - Deduplica entidades por chave_entidade dentro do próprio Excel.
    """
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    log = {"sheets_ok": [], "sheets_ign": [], "linhas_processadas": 0}

    entidades_map = {}   # chave → dict entidade
    dirigentes = []      # lista de dicts

    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for sheet in xl.sheet_names:
        if norm(sheet) in SHEETS_IGNORAR:
            log["sheets_ign"].append(f"{sheet} (resumo)")
            continue
        try:
            hr, score = _detectar_header(xl, sheet)
            if score < 2:
                log["sheets_ign"].append(f"{sheet} (sem cabeçalho reconhecível)")
                continue
            df = xl.parse(sheet, header=hr)
            df.columns = [str(c).strip() for c in df.columns]
            m = _detectar_colunas(df)
            if "designacao" not in m:
                log["sheets_ign"].append(f"{sheet} (sem coluna de entidade)")
                continue

            cat_id = cat_por_sheet(sheet)
            n_add = 0
            for _, row in df.iterrows():
                desig = limpar_str(row.get(m.get("designacao"), ""))
                if len(desig) < 3: continue
                sigla = limpar_str(row.get(m.get("sigla",""), "")) if "sigla" in m else ""
                sioe  = limpar_str(row.get(m.get("sioe_code",""), "")) if "sioe_code" in m else ""
                nif   = limpar_str(row.get(m.get("nif",""), "")) if "nif" in m else ""

                k = chave_entidade(desig, sigla, sioe, nif)
                if k not in entidades_map:
                    eid = slugify(f"{sigla}-{desig}" if sigla else desig)
                    # garantir unicidade do eid
                    base_eid = eid
                    sfx = 2
                    while eid in {e["entity_id"] for e in entidades_map.values()}:
                        eid = f"{base_eid}-{sfx}"
                        sfx += 1
                    entidades_map[k] = {
                        "entity_id":    eid,
                        "designacao":   desig,
                        "siglas":       sigla,
                        "sioe_code":    sioe,
                        "nif":          nif,
                        "categoria_id": cat_id,
                        "ministerio":   limpar_str(row.get(m.get("ministerio",""), "")),
                        "tutela":       "",
                        "website":      limpar_str(row.get(m.get("website",""), "")),
                        "morada":       limpar_str(row.get(m.get("morada",""), "")),
                        "distrito":     limpar_str(row.get(m.get("distrito",""), "")),
                        "estado":       "ativa",
                        "criado_em":    agora,
                        "alterado_em":  agora,
                        "notas":        "",
                    }
                ent = entidades_map[k]
                # enriquecer entidade se houver campos em falta
                for f_csv, f_col in [("siglas","sigla"), ("sioe_code","sioe_code"),
                                     ("nif","nif"), ("ministerio","ministerio"),
                                     ("website","website"), ("morada","morada"),
                                     ("distrito","distrito")]:
                    if not ent.get(f_csv) and f_col in m:
                        v = limpar_str(row.get(m[f_col], ""))
                        if v: ent[f_csv] = v

                # dirigente
                nome   = limpar_str(row.get(m.get("nome_dirigente",""), "")) if "nome_dirigente" in m else ""
                cargo  = limpar_str(row.get(m.get("cargo",""), "")) if "cargo" in m else ""
                emai   = email_ok(row.get(m.get("email",""), ""))     if "email" in m else ""
                tele   = tel_ok(row.get(m.get("telefone",""), ""))    if "telefone" in m else ""
                if nome or emai or cargo:
                    dirigentes.append({
                        "dirig_id":    str(uuid.uuid4())[:12],
                        "entity_id":   ent["entity_id"],
                        "nome":        nome,
                        "cargo":       cargo,
                        "email":       emai,
                        "telefone":    tele,
                        "inicio":      "",
                        "fim":         "",
                        "fonte":       f"INA:{sheet}",
                        "confianca":   "alta",
                        "criado_em":   agora,
                        "alterado_em": agora,
                        "notas":       "",
                    })
                n_add += 1
            log["sheets_ok"].append(f"{sheet}: {n_add} linhas")
            log["linhas_processadas"] += n_add
        except Exception as e:
            log["sheets_ign"].append(f"{sheet} (erro: {e})")
            continue

    # deduplicar dirigentes (mesmo cargo+nome+email na mesma entidade)
    df_d = pd.DataFrame(dirigentes, columns=COLS_DIRIGENTE) if dirigentes else pd.DataFrame(columns=COLS_DIRIGENTE)
    if len(df_d):
        df_d["_k"] = df_d.apply(lambda r: chave_dirigente(r["entity_id"], r["cargo"], r["email"], r["nome"]), axis=1)
        df_d = df_d.drop_duplicates(subset=["_k"]).drop(columns=["_k"])

    df_e = pd.DataFrame(list(entidades_map.values()), columns=COLS_ENTIDADE)
    log["entidades_distintas"] = len(df_e)
    log["dirigentes_distintos"] = len(df_d)
    return df_e, df_d, log

# ═══════════════════════════════════════════════════
# IMPORT SIOE (manual)
# ═══════════════════════════════════════════════════
def importar_sioe(file_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    df = pd.read_excel(io.BytesIO(file_bytes), header=0)
    df.columns = [str(c).strip() for c in df.columns]
    log = {"linhas": len(df)}
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

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

    desig = gc("Designação")
    sigla = gc("Sigla")
    tipo  = gc("Tipo de Entidade")
    sioe  = gc("Código SIOE") if "Código SIOE" in df.columns else pd.Series("", index=df.index)
    mini  = gc("Ministério/Secretaria Regional")

    entidades_map = {}
    dirigentes = []
    for i in range(len(df)):
        d = limpar_str(desig.iloc[i])
        if len(d) < 3: continue
        s  = limpar_str(sigla.iloc[i])
        sc = limpar_str(sioe.iloc[i])
        k  = chave_entidade(d, s, sc, "")
        if k not in entidades_map:
            eid = slugify(f"{s}-{d}" if s else d)
            base_eid = eid; sfx = 2
            while eid in {e["entity_id"] for e in entidades_map.values()}:
                eid = f"{base_eid}-{sfx}"; sfx += 1
            entidades_map[k] = {
                "entity_id": eid, "designacao": d, "siglas": s, "sioe_code": sc, "nif": "",
                "categoria_id": cat_por_tipo_sioe(tipo.iloc[i]),
                "ministerio": limpar_str(mini.iloc[i]), "tutela": "",
                "website": "", "morada": "", "distrito": "",
                "estado": "ativa", "criado_em": agora, "alterado_em": agora, "notas": "",
            }
        ent = entidades_map[k]
        nome  = limpar_str(nome_s.iloc[i])
        cargo = limpar_str(cargo_s.iloc[i])
        emai  = email_ok(email_s.iloc[i])
        tele  = tel_ok(tel_s.iloc[i])
        if nome or emai:
            dirigentes.append({
                "dirig_id": str(uuid.uuid4())[:12], "entity_id": ent["entity_id"],
                "nome": nome, "cargo": cargo, "email": emai, "telefone": tele,
                "inicio": "", "fim": "", "fonte": "SIOE", "confianca": "alta",
                "criado_em": agora, "alterado_em": agora, "notas": "",
            })

    df_e = pd.DataFrame(list(entidades_map.values()), columns=COLS_ENTIDADE)
    df_d = pd.DataFrame(dirigentes, columns=COLS_DIRIGENTE) if dirigentes else pd.DataFrame(columns=COLS_DIRIGENTE)
    if len(df_d):
        df_d["_k"] = df_d.apply(lambda r: chave_dirigente(r["entity_id"], r["cargo"], r["email"], r["nome"]), axis=1)
        df_d = df_d.drop_duplicates(subset=["_k"]).drop(columns=["_k"])
    log["entidades"]  = len(df_e)
    log["dirigentes"] = len(df_d)
    return df_e, df_d, log

# ═══════════════════════════════════════════════════
# MERGE CONTROLADO (com histórico de dirigentes)
# ═══════════════════════════════════════════════════
def merge_entidades(base: pd.DataFrame, novo: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Une base + novo. Devolve (df_final, relatorio).
    Regras: entidades com mesma chave_entidade são fundidas; campos em falta
    na base são preenchidos pelo novo; campos preenchidos não são sobrescritos.
    """
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if len(base) == 0:
        return novo.copy(), {"adicionadas": len(novo), "atualizadas": 0}
    base = base.copy()
    # mapa chave → indice
    def kf(r): return chave_entidade(r.get("designacao",""), r.get("siglas",""), r.get("sioe_code",""), r.get("nif",""))
    base["_k"] = base.apply(kf, axis=1)
    existing = {k: i for i, k in zip(base.index, base["_k"])}

    n_add, n_upd = 0, 0
    novos_rows = []
    for _, r in novo.iterrows():
        k = kf(r)
        if k in existing:
            idx = existing[k]
            for c in COLS_ENTIDADE:
                if c in ("entity_id","criado_em"): continue
                cur = base.at[idx, c] if c in base.columns else ""
                new = r.get(c, "")
                if (not cur or str(cur).strip() in ("","nan")) and (new and str(new).strip()):
                    base.at[idx, c] = new
                    n_upd += 1
            base.at[idx, "alterado_em"] = agora
        else:
            novos_rows.append(r[COLS_ENTIDADE])
            n_add += 1
    if novos_rows:
        base = pd.concat([base.drop(columns=["_k"]), pd.DataFrame(novos_rows, columns=COLS_ENTIDADE)], ignore_index=True)
    else:
        base = base.drop(columns=["_k"])
    return base, {"adicionadas": n_add, "atualizadas": n_upd}

def merge_dirigentes(base: pd.DataFrame, novo: pd.DataFrame, base_entidades: pd.DataFrame,
                     novo_entidades: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Funde dirigentes. Regra fundamental: se para uma (entity_id, cargo) já existe
    um dirigente ATIVO (fim vazio) e chega um novo com nome DIFERENTE, o antigo
    é marcado como histórico (fim=hoje) e o novo é inserido como ativo.
    """
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    hoje  = datetime.now().strftime("%Y-%m-%d")

    # remapear entity_ids do novo para os entity_ids da base
    if len(base_entidades):
        def kf(r): return chave_entidade(r.get("designacao",""), r.get("siglas",""), r.get("sioe_code",""), r.get("nif",""))
        base_entidades = base_entidades.copy()
        base_entidades["_k"] = base_entidades.apply(kf, axis=1)
        novo_entidades = novo_entidades.copy()
        novo_entidades["_k"] = novo_entidades.apply(kf, axis=1)
        mapa = {}
        base_k2id = dict(zip(base_entidades["_k"], base_entidades["entity_id"]))
        for _, r in novo_entidades.iterrows():
            if r["_k"] in base_k2id:
                mapa[r["entity_id"]] = base_k2id[r["_k"]]
        if mapa and len(novo):
            novo = novo.copy()
            novo["entity_id"] = novo["entity_id"].map(lambda x: mapa.get(x, x))

    if len(base) == 0:
        return novo.copy(), {"adicionados": len(novo), "sucessoes": 0, "duplicados": 0}

    base = base.copy()
    n_add, n_suc, n_dup = 0, 0, 0
    for _, r in novo.iterrows():
        eid = r["entity_id"]
        cargo = r.get("cargo","")
        nome  = r.get("nome","")
        email = r.get("email","")
        # Duplicado exacto?
        mask = (
            (base["entity_id"] == eid) &
            (base["cargo"].fillna("").str.lower().str.strip() == str(cargo).lower().strip()) &
            (base["nome"].fillna("").str.lower().str.strip()  == str(nome).lower().strip())
        )
        if mask.any():
            n_dup += 1
            # atualizar email/telefone se em falta
            idx = base.index[mask][0]
            for c in ("email","telefone"):
                if not str(base.at[idx, c] or "").strip() and r.get(c,""):
                    base.at[idx, c] = r[c]
            continue
        # Sucessão: mesmo entity+cargo, nome diferente, e há ativo
        if cargo:
            ativo_mask = (
                (base["entity_id"] == eid) &
                (base["cargo"].fillna("").str.lower().str.strip() == str(cargo).lower().strip()) &
                (base["fim"].fillna("") == "")
            )
            if ativo_mask.any() and nome:
                for idx in base.index[ativo_mask]:
                    base.at[idx, "fim"] = hoje
                    base.at[idx, "alterado_em"] = agora
                    base.at[idx, "notas"] = (str(base.at[idx, "notas"] or "") + " | sucedido").strip(" |")
                n_suc += 1
        # inserir novo
        novo_row = {c: r.get(c,"") for c in COLS_DIRIGENTE}
        if not novo_row.get("dirig_id"): novo_row["dirig_id"] = str(uuid.uuid4())[:12]
        if not novo_row.get("criado_em"): novo_row["criado_em"] = agora
        novo_row["alterado_em"] = agora
        base = pd.concat([base, pd.DataFrame([novo_row], columns=COLS_DIRIGENTE)], ignore_index=True)
        n_add += 1

    return base, {"adicionados": n_add, "sucessoes": n_suc, "duplicados": n_dup}
