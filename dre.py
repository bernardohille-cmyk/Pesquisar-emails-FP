"""
dre.py — Scraping ligeiro do Diário da República.
Devolve SUGESTÕES (não insere nada). A equipa aprova/rejeita na aba 'Revisão'.
Nunca levanta exceções: em caso de falha, devolve lista vazia.
"""
from __future__ import annotations
import json, re, uuid
from datetime import datetime, timedelta
import pandas as pd
import requests

from db import norm, gh_get, gh_put, PATH_PENDENTES

PALAVRAS_CHAVE = [
    "nomeacao", "exoneracao", "designacao",
    "cessacao de funcoes", "tomada de posse", "investidura",
]

def _sid() -> str:
    return str(uuid.uuid4())[:12]

def carregar_pendentes() -> list[dict]:
    try:
        raw, _ = gh_get(PATH_PENDENTES)
        if raw is None:
            return []
        return json.loads(raw.decode("utf-8")).get("pendentes", [])
    except Exception:
        return []

def guardar_pendentes(lista: list[dict]) -> tuple[bool, str]:
    try:
        _, sha = gh_get(PATH_PENDENTES)
        payload = {
            "atualizado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "pendentes": lista,
        }
        ok, msg, _ = gh_put(
            PATH_PENDENTES,
            json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"),
            sha,
            f"Pendentes: {len(lista)} sugestoes",
        )
        return ok, msg
    except Exception as e:
        return False, f"erro guardar_pendentes: {e}"

def buscar_dre(dias: int = 30) -> list[dict]:
    """Pesquisa simples no DRE. Nunca lança exceção."""
    try:
        inicio = (datetime.now() - timedelta(days=dias)).strftime("%Y-%m-%d")
        fim = datetime.now().strftime("%Y-%m-%d")
        url = "https://dre.pt/dre/pesquisa-avancada"
        params = {
            "_PesquisaAvancada_dataPublicacaoInicio": inicio,
            "_PesquisaAvancada_dataPublicacaoFim": fim,
            "search_text": "nomeação OR exoneração OR designação",
            "perPage": 50,
        }
        r = requests.get(
            url, params=params, timeout=20,
            headers={"User-Agent": "AP-Contactos-INA/1.0"},
        )
        if r.status_code != 200:
            return []
        html = r.text
        matches = re.findall(
            r'<a[^>]+href="(/[^"]+detalhe[^"]+)"[^>]*>([^<]{20,300})</a>',
            html, re.IGNORECASE,
        )
        out = []
        for href, titulo in matches[:50]:
            tnorm = norm(titulo)
            kws = [k for k in PALAVRAS_CHAVE if k in tnorm]
            if not kws:
                continue
            out.append({
                "titulo": titulo.strip(),
                "url": "https://dre.pt" + href,
                "palavras": kws,
                "data": fim,
            })
        return out
    except Exception:
        return []

def sugestoes_a_partir_de_dre(entidades_df: pd.DataFrame, dias: int = 30) -> list[dict]:
    """Nunca lança exceção. Devolve lista de sugestões com entity_id casado."""
    try:
        achados = buscar_dre(dias=dias)
        if not achados or len(entidades_df) == 0:
            return []

        desigs = {norm(d): eid for d, eid in
                  zip(entidades_df["designacao"], entidades_df["entity_id"])}
        siglas = {}
        for s, eid in zip(entidades_df["siglas"], entidades_df["entity_id"]):
            for sig in str(s).split(","):
                sig = norm(sig.strip())
                if len(sig) >= 2:
                    siglas[sig] = eid

        sugestoes = []
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for hit in achados:
            tnorm = norm(hit["titulo"])
            ent_match = []
            for d, eid in desigs.items():
                if len(d) >= 8 and d in tnorm:
                    ent_match.append((eid, d, "designacao"))
            for s, eid in siglas.items():
                # FIX: \b em regex normal, não duplo-escapado
                if re.search(r"\b" + re.escape(s) + r"\b", tnorm):
                    ent_match.append((eid, s, "sigla"))
            if not ent_match:
                continue

            if "exoner" in tnorm or "cessac" in tnorm:
                tipo_acao = "exoneracao"
            elif "nomea" in tnorm or "designa" in tnorm or "posse" in tnorm:
                tipo_acao = "nomeacao"
            else:
                tipo_acao = "alteracao"

            vistos = set()
            for eid, evidencia, tipo in ent_match[:3]:
                if eid in vistos:
                    continue
                vistos.add(eid)
                sugestoes.append({
                    "id": _sid(),
                    "criada_em": agora,
                    "fonte": "DRE",
                    "tipo_acao": tipo_acao,
                    "entity_id": eid,
                    "evidencia": f"{tipo}={evidencia}",
                    "titulo": hit["titulo"][:300],
                    "url": hit["url"],
                    "palavras_kw": ",".join(hit["palavras"]),
                    "estado": "pendente",
                })
        return sugestoes
    except Exception:
        return []
