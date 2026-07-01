# Documentação de Campos — API PACTO
# Dashboard Executivo Greenlife

Fonte: HTML da documentação Swagger em api-docs.pactosolucoes.com.br
Atualizado em: 2026-06-15

---

## Configuração da API

| Item | Valor |
|------|-------|
| URL base | `https://apigw.pactosolucoes.com.br` |
| Documentação | `https://api-docs.pactosolucoes.com.br` |
| Autenticação | `Authorization: Bearer <token>` |
| Header obrigatório | `empresaId: <integer>` |
| Content-Type (POST) | `application/json` |

---

## Envelope de Resposta (padrão de todos os endpoints)

```json
{
  "meta": {
    "statusCode": 200,
    "error": null,
    "message": null,
    "messageID": null,
    "messageValue": null
  },
  "content": <object ou array>,
  "totalElements": 150,
  "totalPages": 15,
  "first": true,
  "last": false,
  "numberOfElements": 10,
  "size": 10,
  "number": 0
}
```

| Campo | Tipo | Descrição |
|-------|------|-----------|
| `meta.statusCode` | int32 | Código HTTP da resposta |
| `meta.error` | string | Código do erro (ex: `VALIDATION_ERROR`) |
| `meta.message` | string | Mensagem de erro legível |
| `content` | object/array | Dados da resposta (campos variam por endpoint) |
| `totalElements` | int64 | Total de registros disponíveis |
| `totalPages` | int64 | Total de páginas |
| `first` | boolean | Se é a primeira página |
| `last` | boolean | Se é a última página |
| `numberOfElements` | int64 | Registros nesta página |
| `size` | int64 | Tamanho da página configurado |
| `number` | int64 | Número da página atual (começa em 0) |

---

## Parâmetros Globais Identificados

| Nome | Localização | Tipo | Descrição |
|------|------------|------|-----------|
| `empresaId` | header | integer (int32) | **Obrigatório.** Código da empresa no PACTO |
| `filters` | query | string | Filtros adicionais (formato a confirmar) |
| `page` | query | integer | Número da página (começa em 0) |
| `size` | query | integer | Itens por página |

---

## Endpoints de BI Administrativo (módulo do Dashboard)

> Descrição: "Operações para visualização de Business Intelligence do módulo Administrativo"

| Método | Endpoint | Descrição | Schema `content` |
|--------|----------|-----------|-----------------|
| GET | `/inadimplencia` | Dados de inadimplência por indicador | A confirmar |
| GET | `/gestao-acesso` | Dados de gestão de acessos por indicador | A confirmar |
| GET | `/ltv` | Lifetime Value (LTV) | A confirmar |
| GET | `/churn-prediction` | Dados de churn prediction | A confirmar |
| GET | `/aulas-experimentais` | Dados de aulas experimentais | A confirmar |
| GET | `/meta-financeira/empresa-mes-ano` | Metas financeiras por empresa/mês/ano | A confirmar |
| POST | `/indice-conversao-vendas` | Índice de conversão de vendas por indicador | A confirmar |
| GET | `/movimentacao-contrato` | Movimentação de contratos | A confirmar |
| GET | `/pendencia-clientes` | Pendências de clientes | A confirmar |
| GET | `/observacao-operacao/parcelas-canceladas` | Parcelas canceladas com totais | A confirmar |
| GET | `/observacao-operacao/parcelas-canceladas-sem-totais` | Parcelas canceladas sem totais | A confirmar |
| GET | `/justificativa-operacao/contratos-cancelados-transferidos` | Contratos cancelados/transferidos | A confirmar |
| GET | `/estorno-observacao/admin` | Estornos (administradores) | A confirmar |
| GET | `/estorno-observacao/recorrencia` | Estornos de recorrência | A confirmar |
| GET | `/estorno-observacao/usuario-comum` | Estornos (usuário comum) | A confirmar |
| GET | `/cobranca-por-convenio` | Cobranças por convênio | A confirmar |
| GET | `/clientes/clientes-para-verificar` | Clientes para verificar | A confirmar |
| POST | `/clientes/dados-churn` | Cadastrar dados de Churn | A confirmar |

---

## Endpoint Escolhido para Validação do Fluxo

**`GET /inadimplencia`**

Motivo da escolha:
- Método GET (mais simples)
- Nome autoexplicativo, indicador relevante para dashboard
- Sem parâmetros obrigatórios de body
- Retorna dado estratégico de alto valor

Chamada esperada:
```
GET https://apigw.pactosolucoes.com.br/inadimplencia
Headers:
  Authorization: Bearer <TOKEN>
  empresaId: <EMPRESA_ID>
```

---

## Outros Endpoints Catalogados (249 total)

Módulos identificados além do BI Administrativo:

| Módulo (prefixo) | Quantidade | Uso |
|-----------------|-----------|-----|
| `/cliente/` | ~80 | App / operações de cliente |
| `/psec/alunos/` | ~25 | Personal trainer / avaliações |
| `/clientes/` | ~15 | Consultas de cliente |
| `/v1/cliente`, `/v2/cliente` | ~10 | API REST padrão de cliente |
| `/cadastro-cliente/` | ~8 | Cadastro |
| `/cliente-observacao/` | ~7 | Observações |
| `/contratos/` | ~5 | Contratos |
| `/rel-clientes-geral/` | ~10 | Relatórios gerais |
| `/importacao/` | ~6 | Importação em lote |
| `/psec/clientes/` | ~5 | Clientes no PSEC |

### Endpoints de Relatório de Clientes Geral (relevantes para dashboard)

| Método | Endpoint | Parâmetros documentados |
|--------|----------|------------------------|
| GET | `/rel-clientes-geral/planos` | `filters` (query), `empresaId` (header) |
| GET | `/rel-clientes-geral/pacotes` | A confirmar |
| GET | `/rel-clientes-geral/consultar` | A confirmar |
| GET | `/rel-clientes-geral/categorias` | A confirmar |
| GET | `/rel-clientes-geral/consultores` | A confirmar |
| GET | `/rel-clientes-geral/convenios-desconto` | A confirmar |
| GET | `/rel-clientes-geral/duracoes` | A confirmar |
| GET | `/rel-clientes-geral/eventos` | A confirmar |
| GET | `/rel-clientes-geral/modalidades` | A confirmar |
| GET | `/rel-clientes-geral/origens-sistema` | A confirmar |
| GET | `/rel-clientes-geral/professores` | A confirmar |
