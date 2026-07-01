import json
import csv
import os
import requests
from datetime import datetime, timezone
from calendar import monthrange

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")
BASE_URL = "https://graph.facebook.com/v21.0"


def load_config():
    with open(CONFIG_FILE, encoding="utf-8") as f:
        return json.load(f)


def api_get(path, params, silent=False):
    r = requests.get(f"{BASE_URL}/{path}", params=params, timeout=30)
    if not r.ok and not silent:
        try:
            print(f"    [API erro] {r.json().get('error', {}).get('message', r.text)}")
        except Exception:
            pass
    return r.json() if r.ok else {}


def month_windows(year, month):
    """Divide o mês em janelas de 28 dias (limite da API)."""
    tz = timezone.utc
    _, last_day = monthrange(year, month)
    windows, day = [], 1
    while day <= last_day:
        end = min(day + 27, last_day)
        since = int(datetime(year, month, day, 0, 0, 0, tzinfo=tz).timestamp())
        until = int(datetime(year, month, end, 23, 59, 59, tzinfo=tz).timestamp())
        windows.append((since, until))
        day = end + 1
    return windows


# ─────────────────────────────────────────────────────────────
# BLOCO 1 — Seguidores
# ─────────────────────────────────────────────────────────────

def get_seguidores(ig_id, token):
    """Total de seguidores atual."""
    data = api_get(ig_id, {"fields": "followers_count,username", "access_token": token})
    return data.get("followers_count"), data.get("username", "")


def get_seguidores_ganhos(ig_id, token, year, month):
    """Novos seguidores dia a dia (follower_count diário), somados no mês."""
    ganhos = 0
    for since, until in month_windows(year, month):
        data = api_get(f"{ig_id}/insights", {
            "metric": "follower_count",
            "period": "day",
            "since": since,
            "until": until,
            "access_token": token,
        }, silent=True)
        for item in data.get("data", []):
            if item.get("name") == "follower_count":
                for v in item.get("values", []):
                    val = v.get("value", 0) or 0
                    if val > 0:
                        ganhos += val
    return ganhos


# ─────────────────────────────────────────────────────────────
# BLOCO 2 — Interações
# ─────────────────────────────────────────────────────────────

def get_interacoes_perfil(ig_id, token, year, month):
    """
    Métricas de interação no nível de perfil.
    Retorna total_interactions (total geral) e saves (salvamentos de todos os conteúdos).
    """
    totals = {"total_interactions": 0, "saves": 0}
    for since, until in month_windows(year, month):
        data = api_get(f"{ig_id}/insights", {
            "metric": "total_interactions,saves",
            "period": "day",
            "metric_type": "total_value",
            "since": since,
            "until": until,
            "access_token": token,
        }, silent=True)
        for item in data.get("data", []):
            name = item.get("name")
            val = item.get("total_value", {}).get("value")
            if name in totals and val is not None:
                totals[name] += val
    return totals


def get_interacoes_posts(ig_id, token, year, month):
    """
    Soma curtidas, comentários e compartilhamentos de cada post do período.
    Salvamentos vêm do nível de perfil (mais preciso).
    """
    tz = timezone.utc
    _, last_day = monthrange(year, month)
    since = int(datetime(year, month, 1, 0, 0, 0, tzinfo=tz).timestamp())
    until = int(datetime(year, month, last_day, 23, 59, 59, tzinfo=tz).timestamp())

    data = api_get(f"{ig_id}/media", {
        "fields": "id,timestamp,like_count,comments_count",
        "since": since,
        "until": until,
        "limit": 100,
        "access_token": token,
    })
    posts = data.get("data", [])

    curtidas = comentarios = compartilhamentos = 0
    for post in posts:
        curtidas += post.get("like_count", 0) or 0
        comentarios += post.get("comments_count", 0) or 0
        ins = api_get(f"{post['id']}/insights", {"metric": "shares", "access_token": token}, silent=True)
        for m in ins.get("data", []):
            if m.get("name") == "shares":
                val = (m.get("values") or [{}])[0].get("value") or m.get("value") or 0
                compartilhamentos += val

    return curtidas, comentarios, compartilhamentos, len(posts)


def get_respostas_stories(ig_id, token, year, month):
    """Respostas de stories (funciona apenas para stories ainda ativos)."""
    tz = timezone.utc
    _, last_day = monthrange(year, month)
    since = int(datetime(year, month, 1, 0, 0, 0, tzinfo=tz).timestamp())
    until = int(datetime(year, month, last_day, 23, 59, 59, tzinfo=tz).timestamp())
    replies = 0
    try:
        stories = api_get(f"{ig_id}/stories", {
            "fields": "id", "since": since, "until": until, "access_token": token
        }, silent=True)
        for story in stories.get("data", []):
            ins = api_get(f"{story['id']}/insights", {"metric": "replies", "access_token": token}, silent=True)
            for m in ins.get("data", []):
                if m.get("name") == "replies":
                    replies += m.get("value", 0)
    except Exception:
        pass
    return replies


# ─────────────────────────────────────────────────────────────
# CSV
# ─────────────────────────────────────────────────────────────

FIELDNAMES = [
    # identificação
    "perfil", "mes", "ano", "mes_completo",
    # bloco 1 — seguidores
    "seguidores_total",
    "seguidores_ganhos",
    "seguidores_perdidos",
    "crescimento_liquido",
    # bloco 2 — interações (detalhamento)
    "curtidas",
    "comentarios",
    "compartilhamentos",
    "salvamentos",
    "respostas_stories",
    # bloco 2 — totais
    "interacoes_qualificadas",
    "total_interacoes_api",
    # contexto
    "posts_no_periodo",
    "coletado_em",
]


def load_csv(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8-sig") as f:
        return [r for r in csv.DictReader(f) if r.get("perfil")]


def save_csv(path, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


# ─────────────────────────────────────────────────────────────
# COLETA PRINCIPAL
# ─────────────────────────────────────────────────────────────

def collect_month(year, month, config):
    token = config["access_token"]
    output_path = os.path.join(os.path.dirname(__file__), config["output_csv"])
    existing = load_csv(output_path)
    mes_str = f"{month:02d}"
    ano_str = str(year)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    today = datetime.today()
    _, last_day = monthrange(year, month)
    mes_completo = "sim" if (year < today.year or (year == today.year and month < today.month)) else "nao"

    new_rows = []
    for profile in config["profiles"]:
        name = profile["name"]
        ig_id = profile["ig_user_id"]

        print(f"  [{name}] seguidores...", end=" ", flush=True)
        seg_total, _ = get_seguidores(ig_id, token)

        print("ganhos...", end=" ", flush=True)
        seg_ganhos = get_seguidores_ganhos(ig_id, token, year, month)

        print("interações perfil...", end=" ", flush=True)
        perf = get_interacoes_perfil(ig_id, token, year, month)

        print("posts...", end=" ", flush=True)
        curtidas, comentarios, compartilhamentos, n_posts = get_interacoes_posts(ig_id, token, year, month)

        print("stories...", end=" ", flush=True)
        respostas = get_respostas_stories(ig_id, token, year, month)

        salvamentos = perf.get("saves") or 0
        total_api = perf.get("total_interactions") or 0
        interacoes_qualificadas = comentarios + compartilhamentos + salvamentos + respostas

        # Crescimento líquido e perdidos calculados via histórico
        existing_clean = [r for r in existing if not (
            r["perfil"] == name and r["mes"] == mes_str and r["ano"] == ano_str
        )]
        prev_m = month - 1 if month > 1 else 12
        prev_y = year if month > 1 else year - 1
        prev = next((r for r in existing_clean
                     if r["perfil"] == name and r["mes"] == f"{prev_m:02d}" and r["ano"] == str(prev_y)), None)

        if prev and prev.get("seguidores_total") and seg_total is not None:
            crescimento_liq = seg_total - int(prev["seguidores_total"])
            seg_perdidos = max(0, seg_ganhos - crescimento_liq)
        else:
            crescimento_liq = ""
            seg_perdidos = ""

        print("OK")
        existing = existing_clean
        new_rows.append({
            "perfil": name,
            "mes": mes_str,
            "ano": ano_str,
            "mes_completo": mes_completo,
            "seguidores_total": seg_total if seg_total is not None else "",
            "seguidores_ganhos": seg_ganhos,
            "seguidores_perdidos": seg_perdidos,
            "crescimento_liquido": crescimento_liq,
            "curtidas": curtidas,
            "comentarios": comentarios,
            "compartilhamentos": compartilhamentos,
            "salvamentos": salvamentos,
            "respostas_stories": respostas,
            "interacoes_qualificadas": interacoes_qualificadas,
            "total_interacoes_api": total_api,
            "posts_no_periodo": n_posts,
            "coletado_em": now_str,
        })

    all_rows = existing + new_rows
    all_rows.sort(key=lambda r: (r["ano"], r["mes"], r["perfil"]))
    save_csv(output_path, all_rows)
    print(f"  Salvo em: {output_path}\n")


# ─────────────────────────────────────────────────────────────
# ENTRADA
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse
    from datetime import timedelta

    parser = argparse.ArgumentParser()
    parser.add_argument("--mes", type=int)
    parser.add_argument("--ano", type=int)
    parser.add_argument("--backfill", type=int, help="Coleta todos os meses do ano informado")
    args = parser.parse_args()

    cfg = load_config()
    today = datetime.today()

    if args.backfill:
        ano = args.backfill
        mes_final = today.month if today.year == ano else 12
        print(f"\nBackfill {ano}: {mes_final} meses\n")
        for m in range(1, mes_final + 1):
            print(f"--- {m:02d}/{ano} ---")
            collect_month(ano, m, cfg)
    else:
        if args.mes and args.ano:
            year, month = args.ano, args.mes
        elif args.mes:
            year, month = today.year, args.mes
        else:
            prev = today.replace(day=1) - timedelta(days=1)
            year, month = prev.year, prev.month
        print(f"\nColetando: {month:02d}/{year}\n")
        collect_month(year, month, cfg)
