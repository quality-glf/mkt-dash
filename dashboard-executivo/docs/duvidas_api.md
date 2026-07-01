# Dúvidas e Informações Ausentes — API PACTO

Registradas em: 2026-06-15

---

## 🔴 BLOQUEANTES (impedem execução)

### 1. Bearer Token
O campo `Authorization: Bearer <token>` é obrigatório em todos os endpoints.
A documentação não informa o valor — precisa ser obtido:
- Via suporte PACTO (token fixo de integração), ou
- Via endpoint de login (ver item 3 abaixo)

**Ação:** Solicitar o token à PACTO ou ao responsável pela integração na Greenlife.

---

### 2. `empresaId` — código da Greenlife no sistema PACTO
Todos os endpoints de BI recebem `empresaId` como header (integer).
Não foi identificado na documentação. Pode ser:
- O código da empresa no módulo administrativo do PACTO
- Diferente por unidade (se sim, precisamos de um por filial)

**Ação:** Verificar no sistema PACTO em qual campo aparece o "código da empresa".

---

## 🟡 INCERTEZAS (afetam qualidade dos dados)

### 3. Autenticação via login ou token fixo?
Existe `POST /cliente/logar` e `GET /psec/validateToken`.
Não está claro se o token é:
- **Fixo** (fornecido pela PACTO, não expira ou expira em prazo longo), ou
- **Dinâmico** (gerado via login, expira em minutos/horas e precisa de refresh)

Se for dinâmico, o cliente precisará implementar fluxo de refresh automático.

---

### 4. Formato do parâmetro `filters`
O parâmetro `filters` (query string, tipo string) aparece em vários endpoints mas a documentação não mostra exemplos de valor.
Pode ser:
- JSON serializado: `?filters={"mes":"2026-04","empresaId":1}`
- Formato chave=valor: `?filters=mes:2026-04,empresaId:1`
- Outro formato proprietário

**Ação:** Confirmar com a PACTO qual é o formato ou testar com uma chamada real.

---

### 5. Parâmetros dos endpoints de BI
Os endpoints de BI Administrativo (`/inadimplencia`, `/gestao-acesso`, `/ltv`, etc.)
não têm parâmetros documentados explicitamente além de `empresaId` (header).
Possíveis parâmetros não documentados:
- Filtro por período (mês/ano)
- Filtro por unidade
- Paginação (`page`, `size`)

**Ação:** Validar na primeira chamada real o que a API aceita e retorna.

---

### 6. Schema de resposta (`content`) não documentado
O campo `content` de todos os endpoints retorna `"<object>"` na documentação —
os campos reais só serão conhecidos após a primeira chamada.

---

### 7. Rate limiting
Não há documentação sobre limites de requisições (req/min, req/dia).
O dashboard fará chamadas periódicas; se houver limite, precisamos de cache.

---

## 🟢 CONFIRMADO

| Item | Valor |
|------|-------|
| URL base | `https://apigw.pactosolucoes.com.br` |
| Autenticação | Bearer Token no header `Authorization` |
| Header obrigatório | `empresaId: <integer>` |
| Paginação | `page` (número da página, começa em 0) + `size` (itens por página) |
| Envelope de resposta | `{ meta, content, totalElements, totalPages, first, last, numberOfElements, size, number }` |
| Endpoints de BI disponíveis | 18 endpoints (ver campos_api.md) |
| Total de endpoints na API | 249 |
