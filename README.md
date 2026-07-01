# Dashboard Greenlife Academias — Documentação

## Visão Geral

Dashboard web que consolida três fontes de dados da rede Greenlife Academias:

| Fonte | Script de coleta | Saída |
|---|---|---|
| Leads CRM (Pacto) | `bi_crm_leads.py` | `historico_leads.xlsx` + `greenlife_scraper/leads_dados.json` |
| Instagram / Meta API | `instagram-metrics/coletar_semanal.py` | `instagram-metrics/base_*.xlsx` |
| Clube de Vantagens | `greenlife_scraper/scraper.py` | `greenlife_scraper/dados_greenlife.xlsx` |

O arquivo `greenlife_scraper/dashboard.html` **é gerado automaticamente** por `greenlife_scraper/gerar_dashboard.py`. Nunca edite o HTML diretamente — as mudanças serão sobrescritas na próxima geração.

---

## Arquitetura de arquivos

```
Projetos Claude/
├── bi_crm_leads.py              ← Coleta leads CRM Pacto (24 unidades)
├── historico_leads.xlsx         ← Histórico acumulado de leads (gerado)
├── README.md                    ← Este arquivo
│
├── greenlife_scraper/
│   ├── gerar_dashboard.py       ← Gera o dashboard.html completo
│   ├── dashboard.html           ← Dashboard final (NÃO editar manualmente)
│   ├── scraper.py               ← Scraper Clube de Vantagens (Playwright)
│   ├── leads_dados.json         ← JSON de leads exportado para o dashboard
│   ├── dados_greenlife.xlsx     ← Dados do Clube de Vantagens
│   └── abrir_dashboard.bat      ← Abre e atualiza o dashboard localmente
│
└── instagram-metrics/
    ├── coletar_semanal.py       ← Coleta métricas Instagram via Meta API
    ├── config.json              ← Token de acesso Meta (⚠️ expira ~60 dias)
    ├── base_seguidores.xlsx
    ├── base_interacoes_total.xlsx
    └── base_interacoes_detalhadas.xlsx
```

---

## Como usar localmente

### Abrindo o dashboard com dados atualizados

Dê um duplo clique em `greenlife_scraper/abrir_dashboard.bat`. Ele executa em sequência:

1. `bi_crm_leads.py` — coleta leads do mês atual e atualiza o Excel + JSON
2. `gerar_dashboard.py` — regenera o `dashboard.html` com todos os dados
3. Abre o `dashboard.html` no navegador

### Rodando os scripts individualmente

```bash
# Atualizar apenas leads (mês atual)
python bi_crm_leads.py

# Regenerar o HTML do dashboard
python greenlife_scraper/gerar_dashboard.py

# Coletar Instagram (semana atual)
python instagram-metrics/coletar_semanal.py

# Coletar Clube de Vantagens (requer Playwright instalado)
python greenlife_scraper/scraper.py
```

---

## Coleta de Leads — bi_crm_leads.py

### Como funciona

- **Primeira execução**: busca todos os meses desde Janeiro/2024 até o mês atual
- **Execuções seguintes**: pula meses já coletados e só rebusca o mês atual
- **Proteção contra falhas**: se o script travar no meio, os dados já coletados estão no Excel e serão reutilizados na próxima execução

### Estrutura da API Pacto

- **Endpoint usuários**: `GET https://apigw.pactosolucoes.com.br/bi-crm/usuarios`
- **Endpoint leads**: `POST https://apigw.pactosolucoes.com.br/bi-crm`
- **Autenticação**: Bearer token único por unidade (campo `api_key` em `UNIDADES`)
- **Timestamps**: milissegundos desde epoch, fuso BRT (UTC-3)
- **Campo relevante**: `resultado[identificadorMeta=="CL"].meta`

### Ponto crítico — Rate Limit HTTP 429

A API Pacto limita requisições por tempo. Com 24 unidades e ~30 meses cada, são ~720 chamadas por execução completa.

**Estratégia adotada:**
- `0.8s` de pausa entre cada mês de uma unidade
- `2s` de pausa entre unidades
- Retry automático em caso de 429: aguarda 60s, 120s ou 180s conforme tentativa (até 3 retries)

Se ainda assim falhar em alguma unidade, o script imprime `ERRO:` e continua para a próxima. Os dados daquela unidade não são perdidos — na próxima execução, o script verifica quais meses estão faltando e só busca esses.

### Saídas

**`historico_leads.xlsx`** — duas abas:
- `Histórico`: linha por unidade/ano/mês com valor de leads
- `Resumo Mensal`: pivot por região × mês, fácil visualização

**`greenlife_scraper/leads_dados.json`** — consumido pelo dashboard:
```json
{
  "gerado_em": "29/06/2026 10:30",
  "mes_atual": "06/2026",
  "labels": ["01/2024", "02/2024", ...],
  "unidades": { "Aldeota": { "regiao": "Regiao 1", "meses": { "2024-01": 145, ... } } },
  "regioes": { "Regiao 1": { "2024-01": 520, ... } },
  "totais_mes": { "2024-01": 1230, ... }
}
```

---

## Coleta Instagram — coletar_semanal.py

### Como funciona

Usa a **Meta Graph API v21.0** para coletar métricas do período do mês atual até hoje. Salva em três Excel acumulativos (append-only, linhas duplicadas são detectadas e puladas).

**Perfis coletados** (configurados em `instagram-metrics/config.json`):
- Greenlife Academias (conta principal)
- Greenlife Personal
- Greenlife São Paulo
- Greenlife Family

### Ponto crítico — Expiração do Token

O `access_token` em `config.json` é um **Long-Lived Token** da Meta com validade de ~60 dias. Quando expirar, todas as coletas de Instagram vão falhar com erro `Invalid OAuth access token`.

**Como renovar:**
1. Acesse [developers.facebook.com/tools/explorer](https://developers.facebook.com/tools/explorer)
2. Selecione o app, gere um novo User Token com permissões `instagram_basic`, `instagram_manage_insights`, `pages_read_engagement`
3. Converta para Long-Lived Token via: `GET https://graph.facebook.com/oauth/access_token?grant_type=fb_exchange_token&client_id=APP_ID&client_secret=APP_SECRET&fb_exchange_token=SHORT_TOKEN`
4. Substitua o valor de `access_token` em `instagram-metrics/config.json`

**Para publicação no GitHub Actions**: mover o token para uma variável de ambiente `META_ACCESS_TOKEN` e ler em `config.json` com `os.environ.get("META_ACCESS_TOKEN")`.

---

## Coleta Clube de Vantagens — scraper.py

### Como funciona

Usa **Playwright** (automação de browser) para fazer login em `parceirogreenlife.com.br` e extrair dados das academias parceiras. Salva em `dados_greenlife.xlsx`.

**Credenciais de acesso**: `admin_green@gmail.com` / `Greenlife@2025#`

### Ponto crítico — Dependência do Playwright

O scraper requer o Playwright instalado com os browsers:
```bash
pip install playwright
playwright install chromium
```

Para o GitHub Actions, isso está configurado no workflow YAML (ver seção abaixo).

---

## Geração do Dashboard — gerar_dashboard.py

O script lê os três JSONs/Excels e gera o `dashboard.html` completo do zero. Todo o HTML, CSS e JavaScript está embutido neste script Python.

### Fluxo de geração

```
leads_dados.json ──────────────────────────────────────┐
dados_greenlife.xlsx → extração via openpyxl → dict    ├─→ gerar_dashboard() → dashboard.html
instagram-metrics/base_*.xlsx → extração via openpyxl ┘
```

### Atenção

**Nunca edite o `dashboard.html` manualmente.** Qualquer mudança visual (cores, layout, textos) deve ser feita em `gerar_dashboard.py`. Se editar o HTML diretamente, suas mudanças serão perdidas na próxima vez que `gerar_dashboard.py` rodar.

---

## Publicação Web — GitHub Actions + Netlify

### Estratégia

O repositório está (ou vai ser) conectado ao **Netlify**. A cada push na branch `main`, o Netlify publica automaticamente o `dashboard.html`.

Os dados são atualizados 3x por dia via **GitHub Actions** (cron jobs), que:
1. Executam os scripts de coleta
2. Regeneram o `dashboard.html`
3. Fazem commit e push do arquivo atualizado

### Horários de coleta (GitHub Actions usa UTC)

| Horário BRT | Cron UTC |
|---|---|
| 08:00 | `0 11 * * *` |
| 12:00 | `0 15 * * *` |
| 20:00 | `0 23 * * *` |

### Arquivo de workflow (`.github/workflows/atualizar_dashboard.yml`)

```yaml
name: Atualizar Dashboard

on:
  schedule:
    - cron: '0 11 * * *'   # 08h BRT
    - cron: '0 15 * * *'   # 12h BRT
    - cron: '0 23 * * *'   # 20h BRT
  workflow_dispatch:         # permite disparo manual

jobs:
  atualizar:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
        with:
          fetch-depth: 0     # histórico completo para não perder Excel

      - uses: actions/setup-python@v5
        with:
          python-version: '3.12'

      - name: Instalar dependências
        run: |
          pip install openpyxl requests playwright
          playwright install chromium

      - name: Coletar leads CRM
        env:
          # Nenhuma variável necessária — tokens embutidos no script (⚠️ mover para secrets)
        run: python bi_crm_leads.py

      - name: Coletar Instagram
        env:
          META_ACCESS_TOKEN: ${{ secrets.META_ACCESS_TOKEN }}
        run: python instagram-metrics/coletar_semanal.py

      - name: Coletar Clube de Vantagens
        run: python greenlife_scraper/scraper.py

      - name: Gerar dashboard
        run: python greenlife_scraper/gerar_dashboard.py

      - name: Commit e push
        run: |
          git config user.name  "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git add greenlife_scraper/dashboard.html \
                  greenlife_scraper/leads_dados.json \
                  historico_leads.xlsx \
                  instagram-metrics/base_*.xlsx \
                  greenlife_scraper/dados_greenlife.xlsx
          git diff --cached --quiet || git commit -m "chore: atualiza dados $(date -u +%Y-%m-%d\ %H:%M\ UTC)"
          git push
```

### Segredos necessários no GitHub

| Secret | Valor |
|---|---|
| `META_ACCESS_TOKEN` | Long-lived token Meta Graph API |

Os Bearer tokens da API Pacto atualmente estão hardcoded em `bi_crm_leads.py`. Antes de publicar o repositório **publicamente**, mover para secrets GitHub (`PACTO_TOKEN_ALDEOTA`, etc.) ou para um arquivo `.env` ignorado pelo git.

---

## Pontos de atenção e checklist antes de publicar

- [ ] Token Meta (Instagram) expira a cada ~60 dias — renovar e atualizar secret
- [ ] Bearer tokens Pacto estão hardcoded — mover para GitHub Secrets antes de tornar o repo público
- [ ] `historico_leads.xlsx` e os Excel do Instagram precisam estar commitados no repo para o GitHub Actions não perder o histórico a cada execução
- [ ] O scraper do Clube de Vantagens usa Playwright — testar se roda corretamente em `ubuntu-latest` no Actions
- [ ] Configurar o Netlify para publicar a pasta `greenlife_scraper/` como raiz do site
- [ ] Testar o workflow manualmente (`workflow_dispatch`) antes de ativar os crons

---

## Histórico de decisões técnicas

**Por que o dashboard é gerado por Python e não editado manualmente?**
No início do projeto, o HTML era editado diretamente. Quando múltiplos chats do Claude faziam mudanças em paralelo, as edições se sobrepunham e dados eram perdidos. A solução foi centralizar tudo em `gerar_dashboard.py` — qualquer mudança visual fica versionada em Python, não em HTML gerado.

**Por que salvar histórico em Excel e não só JSON?**
O JSON é consumido pelo dashboard mas não é bom para análise humana. O Excel permite que a equipe de gestão abra o arquivo e veja a evolução mensal sem precisar de ferramentas técnicas.

**Por que 3 coletas por dia e não tempo real?**
A API Pacto tem rate limiting agressivo. Coletas muito frequentes geram erro 429. Três coletas diárias (manhã, meio-dia, fim de tarde) são suficientes para manter o dashboard relevante sem sobrecarregar a API.
