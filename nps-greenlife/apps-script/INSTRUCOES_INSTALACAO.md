# Instalação — NPS Greenlife Academias (Apps Script)

## Arquivos deste pacote

| Arquivo | Função |
|---------|--------|
| `01_Setup.gs` | Cria todas as abas com cabeçalhos e dados iniciais |
| `02_Importacao.gs` | Importa Solvis e WeHelp, consolida as bases |
| `03_Utils.gs` | Funções compartilhadas, triggers, menu |
| `04_ClassificacaoIA.gs` | Classificação automática de comentários com Gemini |
| `05_WebApp.gs` | Endpoint JSON consumido pelo dashboard HTML |

---

## Passo a Passo

### 1. Criar a planilha

1. Acesse **Google Drive** → Novo → Planilha do Google
2. Dê o nome: `NPS_Greenlife_Base_Consolidada`
3. No menu superior: **Extensões → Apps Script**

---

### 2. Copiar os scripts

No editor do Apps Script:

1. Clique em **"+ Novo arquivo"** para cada arquivo abaixo
2. Cole o conteúdo de cada `.gs` correspondente:

| Nome no Apps Script | Arquivo a colar |
|---------------------|-----------------|
| `01_Setup` | `01_Setup.gs` |
| `02_Importacao` | `02_Importacao.gs` |
| `03_Utils` | `03_Utils.gs` |
| `04_ClassificacaoIA` | `04_ClassificacaoIA.gs` |
| `05_WebApp` | `05_WebApp.gs` |

> Você pode manter o arquivo padrão `Código.gs` vazio ou deletá-lo.

---

### 3. Rodar o Setup

1. No Apps Script, selecione a função `setupCompleto`
2. Clique em **▶ Executar**
3. Autorize as permissões quando solicitado
4. Aguarde a mensagem de confirmação

Ao finalizar, a planilha terá todas as 9 abas criadas e formatadas.

---

### 4. Configurar a chave Gemini

1. Acesse **aistudio.google.com** → clique em **"Get API Key"** → crie uma chave gratuita
2. Volte para a planilha → aba **Config**
3. Na linha `API_KEY`, cole sua chave na coluna B
4. Confira que `API_PROVIDER` está como `gemini` e `MODELO` como `gemini-1.5-flash`

---

### 5. Verificar permissões de acesso às planilhas fonte

As planilhas Solvis e WeHelp precisam ser acessíveis pela conta Google que executa o script.

- Se você tem acesso de leitura às planilhas via sua conta Google: já está OK
- Se não tiver acesso, solicite permissão de visualização ao dono das planilhas

---

### 6. Rodar a primeira importação

1. Na planilha, use o menu **🏋️ NPS Greenlife → ↓ Importar apenas Solvis**
2. Verifique a aba `Log_Classificacao` — deve aparecer uma linha `OK`
3. Repita para WeHelp: **🏋️ NPS Greenlife → ↓ Importar apenas WeHelp**
4. Execute **⚙ Consolidar bases**

> Se houver erro de mapeamento de colunas, verifique o `Logger.log` no Apps Script
> (menu Executar → Registros de execução) e ajuste os nomes na função `mapearColunasSolvis`.

---

### 7. Classificar com IA (primeira rodada)

1. Menu **🏋️ NPS Greenlife → 🤖 Classificar comentários (IA)**
2. Aguarde — pode levar alguns minutos dependendo do volume
3. Verifique a aba `Base_Classificada_IA` — registros aparecem com cores por criticidade

---

### 8. Configurar triggers automáticos

1. Menu **🏋️ NPS Greenlife → ⏰ Configurar triggers diários**
2. O pipeline rodará automaticamente às **06h e 18h** (horário de Fortaleza)

---

### 9. Publicar o Web App (endpoint JSON para o dashboard)

1. No Apps Script, clique em **"Implantar" → "Nova implantação"**
2. Tipo: **App da Web**
3. Executar como: **Eu (sua conta)**
4. Quem tem acesso: **Qualquer pessoa**
5. Clique em **Implantar** e autorize
6. **Copie a URL gerada** — você vai precisar dela na Fase 5 (dashboard HTML)

> A URL terá o formato:
> `https://script.google.com/macros/s/[ID]/exec`

---

## Verificações após a instalação

| Verificação | Como fazer |
|-------------|-----------|
| Abas criadas | Confirmar as 9 abas na planilha |
| De_Para populado | Aba `De_Para_Unidades` → deve ter ~85 linhas |
| Importação OK | Aba `Log_Classificacao` → linha com status `OK` |
| Classificação OK | Aba `Base_Classificada_IA` → registros com cores |
| Web App ativo | Acessar a URL do Web App + `?acao=ping` no navegador |

---

## Solução de problemas comuns

**Erro "Não foi possível abrir a planilha Solvis"**
→ Verifique se você tem acesso à planilha e se o ID na Config está correto.

**Erro de mapeamento de colunas (índice -1)**
→ O nome de alguma coluna das planilhas Solvis/WeHelp mudou. Abra o Log de execução
  do Apps Script e veja qual índice retornou -1. Ajuste os termos de busca em
  `mapearColunasSolvis` ou `mapearColunasWeHelp`.

**Classificação muito lenta**
→ Normal. Com 100 comentários/dia em lotes de 20, são ~5 chamadas.
  Cada chamada leva ~3-5 segundos. Total: ~30 segundos por rodada.

**Gemini retorna erro 429 (rate limit)**
→ Reduza `LOTE_TAMANHO` na Config de 20 para 10. O script já tem `sleep(1200ms)`
  entre lotes, mas planos gratuitos têm limite de 15 req/min.
