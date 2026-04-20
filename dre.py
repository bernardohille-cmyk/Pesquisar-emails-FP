"""
dre.py — Scraping ligeiro do Diário da República.

Devolve SUGESTÕES (não insere nada).
A equipa aprova/rejeita na aba 'Revisão' da app.
"""
from __future__ import annotations

import json
import re
import uuid
from datetime import datetime, timedelta
from typing import Optional

import pandas as pd
import requests

from db import (
    norm, gh_get, gh_put, PATH_PENDENTES,
    chave_entidade,
)

DRE_API_BASE = "https://dre.pt/dre/api/lex/diplomas/v1"

PALAVRAS_CHAVE = [
    "nomeação", "nomeacao", "exoneração", "exoneracao",
    "designação", "designacao", "cessação de funções", "cessacao de funcoes",
    "tomada de posse", "investidura",
]

def _sugestao_id() -> str:
    return str(uuid.uuid4())[:12]

def carregar_pendentes() -> list[dict]:
    raw, _ = gh_get(PATH_PENDENTES)
    if raw is None: return []
    try:
        return json.loads(raw.decode("utf-8")).get("pendentes", [])
    except Exception:
        return []

def guardar_pendentes(lista: list[dict]) -> tuple[bool, str]:
    raw_atual, sha = gh_get(PATH_PENDENTES)
    payload = {
        "atualizado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "pendentes": lista,
    }
    ok, msg, _ = gh_put(PATH_PENDENTES, json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"),
                        sha, f"Pendentes: {len(lista)} sugestões")
    return ok, msg

def buscar_dre(dias: int = 30) -> list[dict]:
    """
    Pesquisa publicações no DRE com palavras-chave de nomeação/exoneração.
    Devolve lista de dicts com {titulo, sumario, data, url, palavras_chave_match}.
    
    NOTA: O DRE não tem API pública estável. Esta função usa a pesquisa pública e
    pode falhar silenciosamente — nesse caso devolve [].
    """
    resultados = []
    inicio = (datetime.now() - timedelta(days=dias)).strftime("%Y-%m-%d")
    fim    = datetime.now().strftime("%Y-%m-%d")
    try:
        # Endpoint público de pesquisa do DRE (varia; usar com cautela)
        url = "https://dre.pt/dre/pesquisa-avancada"
        params = {
            "_PesquisaAvancada_dataPublicacaoInicio": inicio,
            "_PesquisaAvancada_dataPublicacaoFim": fim,
            "search_text": "nomeação OR exoneração OR designação",
            "perPage": 50,
        }
        r = requests.get(url, params=params, timeout=20,
                         headers={"User-Agent": "AP-Contactos-INA/1.0"})
        if r.status_code != 200:
            return []
        # parsing básico de HTML — extrai títulos e links
        html = r.text
        # padrão genérico: <a href="/...detalhe...">TITULO</a>
        matches = re.findall(
            r'<a[^>]+href="(/[^"]+detalhe[^"]+)"[^>]*>([^<]{20,300})</a>',
            html, re.IGNORECASE)
        for href, titulo in matches[:50]:
            tnorm = norm(titulo)
            kws = [k for k in PALAVRAS_CHAVE if k in tnorm]
            if not kws: continue
            resultados.append({
                "titulo": titulo.strip(),
                "url":    "https://dre.pt" + href,
                "palavras": kws,
                "data":   fim,
            })
    except Exception:
        return []
    return resultados

def sugestoes_a_partir_de_dre(entidades_df: pd.DataFrame, dias: int = 30) -> list[dict]:
    """
    Para cada hit do DRE, tenta encontrar a entidade na nossa base.
    Só gera sugestão se houver match >= 1 entidade conhecida.
    """
    achados = buscar_dre(dias=dias)
    if not achados or len(entidades_df) == 0: return []

    # mapa simples: lower(designação) e siglas → entity_id
    desigs = {norm(d): eid for d, eid in zip(entidades_df["designacao"], entidades_df["entity_id"])}
    siglas = {}
    for s, eid in zip(entidades_df["siglas"], entidades_df["entity_id"]):
        for sig in str(s).split(","):
            sig = norm(sig.strip())
            if len(sig) >= 2: siglas[sig] = eid

    sugestoes = []
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for hit in achados:
        tnorm = norm(hit["titulo"])
        ent_match = []
        for d, eid in desigs.items():
            if len(d) >= 8 and d in tnorm:
                ent_match.append((eid, d, "designacao"))
        for s, eid in siglas.items():
            if re.search(r"\\b" + re.escape(s) + r"\\b", tnorm):
                ent_match.append((eid, s, "sigla"))
        if not ent_match: continue
        for eid, evidencia, tipo in ent_match[:3]:
            tipo_acao = "exoneracao" if "exoner" in tnorm else (
                       "nomeacao"   if "nomea" in tnorm or "designa" in tnorm else "alteracao")
            sugestoes.append({
                "id":           _sugestao_id(),
                "criada_em":    agora,
                "fonte":        "DRE",
                "tipo_acao":    tipo_acao,
                "entity_id":    eid,
                "evidencia":    f"{tipo}={evidencia}",
                "titulo":       hit["titulo"][:300],
                "url":          hit["url"],
                "palavras_kw":  ",".join(hit["palavras"]),
                "estado":       "pendente",
            })
    return sugestoes
