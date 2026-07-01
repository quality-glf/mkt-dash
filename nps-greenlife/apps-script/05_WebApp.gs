// =============================================================================
// NPS GREENLIFE ACADEMIAS — WEB APP (endpoint JSON para o dashboard)
// Arquivo: 05_WebApp.gs
// =============================================================================

/**
 * Endpoint principal. Recebe parâmetros via GET e devolve JSON filtrado.
 *
 * Parâmetros aceitos:
 *   ?acao=dados          → retorna Base_Classificada_IA (com filtros opcionais)
 *   ?acao=meta           → retorna meta NPS e configurações do dashboard
 *   ?acao=unidades       → retorna lista de unidades únicas
 *   ?acao=regioes        → retorna lista de regiões
 *   ?acao=ping           → health check
 *
 * Filtros opcionais para acao=dados:
 *   &mes=04/2026         → filtra por mês/ano
 *   &regiao=Região 1     → filtra por região
 *   &unidade=Aldeota     → filtra por unidade padronizada
 *   &plataforma=Solvis   → filtra por plataforma
 */
function doGet(e) {
  const params = e ? e.parameter : {};
  const acao = params.acao || 'dados';

  try {
    let resultado;

    switch (acao) {
      case 'ping':
        resultado = { status: 'ok', timestamp: new Date().toISOString(), versao: '1.0' };
        break;

      case 'meta':
        resultado = lerMetaDashboard();
        break;

      case 'unidades':
        resultado = lerUnidades();
        break;

      case 'regioes':
        resultado = lerRegioes();
        break;

      case 'pendentes':
        resultado = lerRegistrosPendentes(params);
        break;

      case 'dados':
      default:
        resultado = lerDadosFiltrados(params);
        break;
    }

    return ContentService
      .createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    const erro = { erro: true, mensagem: err.message, timestamp: new Date().toISOString() };
    return ContentService
      .createTextOutput(JSON.stringify(erro))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =============================================================================
// ENDPOINT POST — recebe registros classificados e grava no Sheets
// Body JSON esperado: { "token": "...", "registros": [ {campos...} ] }
// =============================================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // Validação de token simples
    const config = lerConfig();
    const tokenEsperado = config.WEBHOOK_TOKEN || '';
    if (tokenEsperado && body.token !== tokenEsperado) {
      return ContentService
        .createTextOutput(JSON.stringify({ erro: true, mensagem: 'Token inválido.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const registros = body.registros;
    if (!Array.isArray(registros) || registros.length === 0) {
      return ContentService
        .createTextOutput(JSON.stringify({ erro: true, mensagem: 'Nenhum registro recebido.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const gravados = gravarClassificacoes(registros);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, gravados: gravados, timestamp: new Date().toISOString() }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ erro: true, mensagem: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Grava as classificações recebidas em Base_Classificada_IA e marca na Base_Consolidada.
 * Cada registro é um objeto com os campos de classificação + dados originais.
 */
function gravarClassificacoes(registros) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaClassificada = ss.getSheetByName('Base_Classificada_IA');
  const abaConsolidada = ss.getSheetByName('Base_Consolidada');
  const config = lerConfig();
  const agora = Utilities.formatDate(new Date(), 'America/Fortaleza', 'dd/MM/yyyy HH:mm:ss');
  const versaoPrompt = config.VERSAO_PROMPT || 'v1-claude-code';

  // Mapa id -> número da linha real (1-based) para upsert
  const classifData = abaClassificada.getDataRange().getValues();
  const idParaLinha = {};
  for (let i = 1; i < classifData.length; i++) {
    const id = String(classifData[i][0] || '').trim();
    if (id && !id.startsWith('⚠️')) idParaLinha[id] = i + 1;
  }

  const linhasNovas = [];
  const idsNovos = new Set();
  let atualizados = 0;

  for (const r of registros) {
    const id = String(r.id_unico || '').trim();
    if (!id) continue;

    const dataResp = String(r.data_resposta || '');
    const mesAno = String(r.mes_ano || '');
    const temComentario = r.tem_comentario === true || String(r.tem_comentario).toUpperCase() === 'TRUE' ? 'TRUE' : 'FALSE';

    const linhaDados = [
      id,
      r.plataforma || '',
      dataResp,
      mesAno,
      r.dia || '',
      r.unidade_original || '',
      r.unidade_padronizada || '',
      r.regiao || '',
      '',                                                      // clientes_ativos
      r.nota !== undefined ? r.nota : '',
      r.tipo_nps || '',
      '',                                                      // categoria_original
      r.comentario_original || '',
      r.comentario_tratado || '',
      temComentario,
      r.comentario_sem_sentido ? 'TRUE' : 'FALSE',
      r.tema_principal || 'Outro',
      r.tema_secundario || '',
      r.sentimento || 'Neutro',
      r.criticidade || 'Baixa',
      r.tipo_comentario || 'Neutro',
      r.resumo || '',
      r.requer_acao ? 'TRUE' : 'FALSE',
      r.risco_cancelamento ? 'TRUE' : 'FALSE',
      r.palavra_critica || '',
      r.acao_sugerida || '',
      r.prioridade_acao || 'Baixa',
      r.problema_recorrente ? 'TRUE' : 'FALSE',
      r.ocultar_nome_colaborador ? 'TRUE' : 'FALSE',
      agora,
      versaoPrompt + '-CLAUDE-CODE',
    ];

    if (idParaLinha[id]) {
      // Atualiza linha existente (upsert)
      abaClassificada.getRange(idParaLinha[id], 1, 1, linhaDados.length).setValues([linhaDados]);
      atualizados++;
    } else {
      linhasNovas.push(linhaDados);
      idsNovos.add(id);
    }
  }

  // Insere linhas novas em batch
  if (linhasNovas.length > 0) {
    const linhaDois = String(abaClassificada.getRange(2, 1).getValue());
    const linhaInicio = linhaDois.startsWith('⚠️') ? 2 : abaClassificada.getLastRow() + 1;
    abaClassificada.getRange(linhaInicio, 1, linhasNovas.length, linhasNovas[0].length)
      .setValues(linhasNovas);
  }

  // Marca como classificados na Base_Consolidada (apenas IDs novos)
  if (idsNovos.size > 0) {
    const consolidadaData = abaConsolidada.getDataRange().getValues();
    for (let i = 1; i < consolidadaData.length; i++) {
      if (idsNovos.has(String(consolidadaData[i][0] || ''))) {
        abaConsolidada.getRange(i + 1, 25).setValue('TRUE');
      }
    }
  }

  atualizarConfigTimestamp('ULTIMO_PROCESSAMENTO_IA', 'CLAUDE-CODE');
  return linhasNovas.length + atualizados;
}

// =============================================================================
// REGISTROS PENDENTES DE CLASSIFICAÇÃO (para agente externo)
// GET ?acao=pendentes&limite=50
// Retorna registros da Base_Consolidada com tem_comentario=TRUE e não classificados,
// respeitando DATA_INICIO_CLASSIFICACAO do Config.
// =============================================================================
function lerRegistrosPendentes(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaConsolidada = ss.getSheetByName('Base_Consolidada');
  const dados = abaConsolidada.getDataRange().getValues();
  const limite = Math.min(parseInt(params.limite) || 50, 500);

  // IDs já classificados
  const abaClassificada = ss.getSheetByName('Base_Classificada_IA');
  const classifData = abaClassificada.getDataRange().getValues();
  const idsClassificados = new Set();
  for (let i = 1; i < classifData.length; i++) {
    const id = String(classifData[i][0] || '').trim();
    if (id && !id.startsWith('⚠️')) idsClassificados.add(id);
  }

  const dataInicio = lerDataInicioClassificacao();
  const pendentes = [];

  for (let i = 1; i < dados.length && pendentes.length < limite; i++) {
    const linha = dados[i];
    const id = String(linha[0] || '').trim();
    if (!id || id.startsWith('⚠️')) continue;
    if (idsClassificados.has(id)) continue;

    const temComentario = linha[13] === true || String(linha[13]).toUpperCase() === 'TRUE';

    // Filtro de data
    let dataResposta = null;
    if (linha[2] instanceof Date) {
      dataResposta = linha[2];
    } else if (linha[2]) {
      const p = String(linha[2]).split('/');
      if (p.length === 3) dataResposta = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
    }
    if (!dataResposta || dataResposta < dataInicio) continue;

    pendentes.push({
      id_unico:             id,
      plataforma:           String(linha[1] || ''),
      data_resposta:        linha[2] instanceof Date
                              ? Utilities.formatDate(linha[2], 'America/Fortaleza', 'dd/MM/yyyy')
                              : String(linha[2] || ''),
      mes_ano:              linha[3] instanceof Date
                              ? Utilities.formatDate(linha[3], 'America/Fortaleza', 'MM/yyyy')
                              : String(linha[3] || ''),
      dia:                  linha[4],
      unidade_original:     String(linha[6] || ''),
      unidade_padronizada:  String(linha[7] || ''),
      regiao:               String(linha[8] || ''),
      nota:                 Number(linha[9]),
      tipo_nps:             String(linha[10] || ''),
      comentario_original:  String(linha[11] || ''),
      comentario_tratado:   String(linha[12] || ''),
      tem_comentario:       String(linha[13]),
      avaliacao_recepcao:   linha[14],
      avaliacao_professor:  linha[15],
      avaliacao_limpeza:    linha[16],
      avaliacao_aulas:      linha[17],
      avaliacao_equipamentos: linha[18],
      avaliacao_ambiente:   linha[19],
    });
  }

  return {
    pendentes: pendentes,
    total: pendentes.length,
    data_inicio_filtro: Utilities.formatDate(dataInicio, 'America/Fortaleza', 'dd/MM/yyyy'),
    gerado_em: new Date().toISOString(),
  };
}

// =============================================================================
// LEITURA E FILTRAGEM DOS DADOS
// =============================================================================
function lerDadosFiltrados(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('Base_Classificada_IA');
  const dados = aba.getDataRange().getValues();

  if (dados.length < 2) return { registros: [], total: 0 };

  const cabecalho = dados[0];
  const idx = montarIndicesColunas(cabecalho);

  // Filtros
  const filtroMes       = params.mes       || '';
  const filtroRegiao    = params.regiao    || '';
  const filtroUnidade   = params.unidade   || '';
  const filtroPlataforma= params.plataforma|| '';
  const filtroCriticidade= params.criticidade || '';
  const filtroTema      = params.tema      || '';

  const registros = [];

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const id = String(linha[idx.id_unico] || '').trim();
    if (!id || id.startsWith('⚠️')) continue;

    // Aplica filtros
    if (filtroMes && String(linha[idx.mes_ano]) !== filtroMes) continue;
    if (filtroRegiao && String(linha[idx.regiao]) !== filtroRegiao) continue;
    if (filtroUnidade && String(linha[idx.unidade_padronizada]) !== filtroUnidade) continue;
    if (filtroPlataforma && String(linha[idx.plataforma]) !== filtroPlataforma) continue;
    if (filtroCriticidade && String(linha[idx.criticidade_ia]) !== filtroCriticidade) continue;
    if (filtroTema && String(linha[idx.tema_principal_ia]) !== filtroTema) continue;

    // Monta objeto — omite dados sensíveis
    const ocultarNome = String(linha[idx.ocultar_nome_colaborador]) === 'TRUE';
    const comentarioOriginal = String(linha[idx.comentario_original] || '');

    registros.push({
      id:                      id,
      plataforma:              String(linha[idx.plataforma] || ''),
      data_resposta:           String(linha[idx.data_resposta] || ''),
      mes_ano:                 String(linha[idx.mes_ano] || ''),
      dia:                     linha[idx.dia],
      unidade_padronizada:     String(linha[idx.unidade_padronizada] || ''),
      regiao:                  String(linha[idx.regiao] || ''),
      nota:                    Number(linha[idx.nota]),
      tipo_nps:                String(linha[idx.tipo_nps] || ''),
      tem_comentario:          String(linha[idx.tem_comentario]) === 'TRUE',
      comentario_sem_sentido:  String(linha[idx.comentario_sem_sentido]) === 'TRUE',
      comentario_original:     ocultarNome ? ocultarNomesNoTexto(comentarioOriginal) : comentarioOriginal,
      resumo_ia:               String(linha[idx.resumo_ia] || ''),
      tema_principal_ia:       String(linha[idx.tema_principal_ia] || ''),
      tema_secundario_ia:      String(linha[idx.tema_secundario_ia] || ''),
      sentimento_ia:           String(linha[idx.sentimento_ia] || ''),
      criticidade_ia:          String(linha[idx.criticidade_ia] || ''),
      tipo_comentario_ia:      String(linha[idx.tipo_comentario_ia] || ''),
      requer_acao:             String(linha[idx.requer_acao]) === 'TRUE',
      risco_cancelamento:      String(linha[idx.risco_cancelamento]) === 'TRUE',
      palavra_critica:         String(linha[idx.palavra_critica] || ''),
      acao_sugerida_ia:        String(linha[idx.acao_sugerida_ia] || ''),
      prioridade_acao:         String(linha[idx.prioridade_acao] || ''),
      problema_recorrente:     String(linha[idx.problema_recorrente]) === 'TRUE',
    });
  }

  return {
    registros: registros,
    total: registros.length,
    gerado_em: new Date().toISOString(),
    filtros_aplicados: { mes: filtroMes, regiao: filtroRegiao, unidade: filtroUnidade, plataforma: filtroPlataforma },
  };
}

function montarIndicesColunas(cabecalho) {
  const idx = {};
  cabecalho.forEach((nome, i) => { idx[String(nome).trim()] = i; });
  return idx;
}

// =============================================================================
// META E CONFIGURAÇÕES
// =============================================================================
function lerMetaDashboard() {
  const config = lerConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaRegras = ss.getSheetByName('Regras_NPS');
  const regras = abaRegras.getDataRange().getValues();

  const params = {};
  regras.forEach(linha => {
    if (linha[0]) params[String(linha[0]).trim()] = String(linha[1]).trim();
  });

  return {
    meta_nps: Number(params.META_NPS_GERAL) || 60,
    status_acima_meta: Number(params.STATUS_ACIMA_META) || 60,
    status_atencao_min: Number(params.STATUS_ATENCAO_MIN) || 45,
    ultimo_processamento: config.ULTIMO_PROCESSAMENTO_IMPORTACAO || '',
    ultimo_processamento_ia: config.ULTIMO_PROCESSAMENTO_IA || '',
    versao: '1.0',
  };
}

// =============================================================================
// UNIDADES E REGIÕES
// =============================================================================
function lerUnidades() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('De_Para_Unidades');
  const dados = aba.getDataRange().getValues();

  const unidades = [];
  const visto = new Set();

  for (let i = 1; i < dados.length; i++) {
    const pad = String(dados[i][1] || '').trim();
    const regiao = String(dados[i][2] || '').trim();
    const ativa = String(dados[i][4] || '').trim();

    if (pad && pad !== '[EXCLUIR]' && ativa === 'S' && !visto.has(pad)) {
      visto.add(pad);
      unidades.push({ unidade: pad, regiao: regiao });
    }
  }

  return unidades.sort((a, b) => {
    if (a.regiao !== b.regiao) return a.regiao.localeCompare(b.regiao);
    return a.unidade.localeCompare(b.unidade);
  });
}

function lerRegioes() {
  const unidades = lerUnidades();
  const regioes = [...new Set(unidades.map(u => u.regiao))].sort();
  return regioes;
}

// =============================================================================
// OCULTAÇÃO DE NOMES SENSÍVEIS
// =============================================================================
function ocultarNomesNoTexto(texto) {
  // Estratégia: substitui sequências que parecem nomes próprios por [NOME OCULTADO]
  // Padrão: palavra com inicial maiúscula seguida de outra palavra com inicial maiúscula
  return texto.replace(/\b([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÈÌÒÙ][a-záéíóúâêîôûãõàèìòù]+(?:\s+[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÈÌÒÙ][a-záéíóúâêîôûãõàèìòù]+)+)\b/g, '[NOME OCULTADO]');
}
