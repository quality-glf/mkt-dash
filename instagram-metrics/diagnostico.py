import json
import requests
from datetime import datetime, timezone
from calendar import monthrange

CONFIG_FILE = "config.json"
BASE_URL = "https://graph.facebook.com/v21.0"

with open(CONFIG_FILE, encoding="utf-8") as f:
    config = json.load(f)

TOKEN = config["access_token"]
YEAR, MONTH = 2026, 6
UNTIL_DAY = 23  # coleta até dia 23


def api_get(path, params):
    url = f"{BASE_URL}/{path}"
    r = requests.get(url, params={**params, "access_token": TOKEN}, timeout=30)
    full_url = r.request.url.replace(TOKEN, "***TOKEN***")
    return r.json(), full_url


def ts(year, month, day, hour=0, minute=0, second=0):
    return int(datetime(year, month, day, hour, minute, second, tzinfo=timezone.utc).timestamp())


sep = "-" * 70

for profile in config["profiles"]:
    name = profile["name"]
    ig_id = profile["ig_user_id"]

    print(f"\n{'='*70}")
    print(f"  PERFIL: {name}  (ig_user_id: {ig_id})")
    print(f"  Período: 01/06/2026 00:00 UTC -> {UNTIL_DAY:02d}/06/2026 23:59 UTC")
    print(f"{'='*70}")

    # ── 1. SEGUIDORES ATUAIS ──────────────────────────────────────────────
    print(f"\n[1] SEGUIDORES ATUAIS")
    data, url = api_get(ig_id, {"fields": "followers_count,username"})
    print(f"    Endpoint : GET /{ig_id}?fields=followers_count,username")
    print(f"    Resposta : followers_count={data.get('followers_count')}  username={data.get('username')}")
    followers = data.get("followers_count")

    # ── 2. MÉTRICAS DE PERFIL (reach, interactions, etc.) ────────────────
    print(f"\n[2] MÉTRICAS DE PERFIL  (janelas de ≤28 dias)")
    metrics = ["total_interactions", "reach", "follows_and_unfollows",
               "profile_views", "accounts_engaged", "website_clicks"]
    prof_totals = {m: 0 for m in metrics}

    since1, until1 = ts(YEAR, MONTH, 1), ts(YEAR, MONTH, min(28, UNTIL_DAY), 23, 59, 59)
    since2, until2 = ts(YEAR, MONTH, 29), ts(YEAR, MONTH, UNTIL_DAY, 23, 59, 59)
    windows = [(since1, until1, 1, min(28, UNTIL_DAY))]
    if UNTIL_DAY > 28:
        windows.append((since2, until2, 29, UNTIL_DAY))

    for since, until, d_ini, d_fim in windows:
        print(f"\n    Janela: dia {d_ini:02d}/06 -> {d_fim:02d}/06")
        print(f"    Endpoint: GET /{ig_id}/insights")
        print(f"    Params  : metric={','.join(metrics)}")
        print(f"              period=day | metric_type=total_value")
        print(f"              since={since} ({datetime.utcfromtimestamp(since).strftime('%d/%m %H:%M')} UTC)")
        print(f"              until={until} ({datetime.utcfromtimestamp(until).strftime('%d/%m %H:%M')} UTC)")
        data, _ = api_get(f"{ig_id}/insights", {
            "metric": ",".join(metrics),
            "period": "day",
            "metric_type": "total_value",
            "since": since,
            "until": until,
        })
        if "error" in data:
            print(f"    ERRO: {data['error']['message']}")
            continue
        for item in data.get("data", []):
            m = item.get("name")
            val = item.get("total_value", {}).get("value")
            print(f"    {m:30s} = {val}")
            if m in prof_totals and val is not None:
                prof_totals[m] += val

    print(f"\n    TOTAIS ACUMULADOS (soma das janelas):")
    for m, v in prof_totals.items():
        print(f"    {m:30s} = {v}")

    # ── 3. POSTS DO PERÍODO ──────────────────────────────────────────────
    print(f"\n[3] POSTS PUBLICADOS NO PERÍODO")
    since_media = ts(YEAR, MONTH, 1)
    until_media = ts(YEAR, MONTH, UNTIL_DAY, 23, 59, 59)
    print(f"    Endpoint: GET /{ig_id}/media")
    print(f"    Params  : fields=id,timestamp,like_count,comments_count")
    print(f"              since={since_media} -> until={until_media}")
    data, _ = api_get(f"{ig_id}/media", {
        "fields": "id,timestamp,like_count,comments_count",
        "since": since_media,
        "until": until_media,
        "limit": 100,
    })
    posts = data.get("data", [])
    print(f"    Posts encontrados: {len(posts)}")

    total_likes = total_comments = total_shares = total_saves = 0
    for post in posts:
        pid = post["id"]
        ts_post = post.get("timestamp", "")[:10]
        lk = post.get("like_count", 0) or 0
        cm = post.get("comments_count", 0) or 0
        total_likes += lk
        total_comments += cm

        ins_data, _ = api_get(f"{pid}/insights", {"metric": "shares,saved"})
        sh = sv = 0
        for m in ins_data.get("data", []):
            val = (m.get("values") or [{}])[0].get("value") or m.get("value") or 0
            if m.get("name") == "shares": sh = val
            elif m.get("name") == "saved": sv = val
        total_shares += sh
        total_saves += sv
        print(f"    Post {pid} ({ts_post}) | curtidas={lk} comentarios={cm} compartilhamentos={sh} salvamentos={sv}")

    print(f"\n    SOMA DOS POSTS:")
    print(f"    curtidas          = {total_likes}")
    print(f"    comentarios       = {total_comments}")
    print(f"    compartilhamentos = {total_shares}")
    print(f"    salvamentos       = {total_saves}")
    interacoes = total_comments + total_shares + total_saves
    print(f"    interacoes_qualificadas (com+comp+salv) = {interacoes}")

    # ── 4. RESUMO FINAL ──────────────────────────────────────────────────
    print(f"\n{sep}")
    print(f"  RESUMO FINAL — {name} — 01 a {UNTIL_DAY:02d}/06/2026")
    print(f"{sep}")
    print(f"  seguidores_final        = {followers}   [fonte: /{ig_id}?fields=followers_count]")
    print(f"  curtidas                = {total_likes}   [fonte: soma de like_count de {len(posts)} posts via /{ig_id}/media]")
    print(f"  comentarios             = {total_comments}   [fonte: soma de comments_count dos posts]")
    print(f"  compartilhamentos       = {total_shares}   [fonte: insights de cada post, metric=shares]")
    print(f"  salvamentos             = {total_saves}   [fonte: insights de cada post, metric=saved]")
    print(f"  interacoes_qualificadas = {interacoes}   [com + comp + salv]")
    print(f"  reach                   = {prof_totals['reach']}   [fonte: /{ig_id}/insights metric=reach]")
    print(f"  profile_views           = {prof_totals['profile_views']}   [fonte: /{ig_id}/insights metric=profile_views]")
    print(f"  website_clicks          = {prof_totals['website_clicks']}   [fonte: /{ig_id}/insights metric=website_clicks]")
    print(f"  total_interactions_api  = {prof_totals['total_interactions']}   [fonte: /{ig_id}/insights metric=total_interactions]")
    print(f"  accounts_engaged        = {prof_totals['accounts_engaged']}   [fonte: /{ig_id}/insights metric=accounts_engaged]")
    print(f"  follows_and_unfollows   = {prof_totals['follows_and_unfollows']}   [fonte: /{ig_id}/insights metric=follows_and_unfollows]")
    print(f"  posts_no_periodo        = {len(posts)}   [fonte: /{ig_id}/media com filtro de data]")
    print()
