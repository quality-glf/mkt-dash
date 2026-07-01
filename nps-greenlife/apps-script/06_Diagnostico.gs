// =============================================================================
// NPS GREENLIFE — DIAGNÓSTICO E CORREÇÃO
// Arquivo: 06_Diagnostico.gs
// Execute diagnosticarBase() para ver o que está acontecendo
// =============================================================================

/**
 * Roda diagnóstico completo e exibe relatório no log.
 * Execute esta função pelo editor do Apps Script.
 */
function diagnosticarBase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaConsolidada = ss.getSheetByName('Base_Consolidada');
  const dados = abaConsolidada.getDataRange().getValues();

  Logger.log('========================================');
  Logger.log('DIAGNÓSTICO — Base_Consolidada');
  Logger.log('Total de linhas (com cabeçalho): ' + dados.length);
  Logger.log('========================================');

  if (dados.length < 2) {
    Logger.log('❌ Base_Consolidada está vazia! Execute a importação primeiro.');
    return;
  }

  // Mostra cabeçalho com índices
  Logger.log('\n--- CABEÇALHO (índice: nome) ---');
  dados[0].forEach((col, i) => Logger.log(i + ': ' + col));

  // Amostra das primeiras 5 linhas de dados
  Logger.log('\n--- AMOSTRA (primeiras 5 linhas de dados) ---');
  const amostra = Math.min(6, dados.length);
  for (let i = 1; i < amostra; i++) {
    const linha = dados[i];
    Logger.log('\nLinha ' + i + ':');
    Logger.log('  [0] id_unico:           ' + linha[0]);
    Logger.log('  [1] plataforma:         ' + linha[1]);
    Logger.log('  [2] data_resposta:      ' + linha[2] + ' (tipo: ' + typeof linha[2] + ')');
    Logger.log('  [3] mes_ano:            ' + linha[3] + ' (tipo: ' + typeof linha[3] + ')');
    Logger.log('  [4] dia:                ' + linha[4]);
    Logger.log('  [5] hora:               ' + linha[5]);
    Logger.log('  [7] unidade_padronizada:' + linha[7]);
    Logger.log('  [8] regiao:             ' + linha[8]);
    Logger.log('  [9] nota:               ' + linha[9]);
    Logger.log('  [10] tipo_nps:          ' + linha[10]);
    Logger.log('  [11] comentario:        ' + String(linha[11]).substring(0, 60));
    Logger.log('  [13] tem_comentario:    ' + linha[13]);
    Logger.log('  [24] _classificado_ia:  ' + linha[24]);
    Logger.log('  dentroDoEscopo:         ' + dentroDoEscopo(String(linha[3])));
  }

  // Contagem por situação
  let totalValidos = 0, semMesAno = 0, foraEscopo = 0, semComentario = 0, comComentario = 0;
  for (let i = 1; i < dados.length; i++) {
    const id = String(dados[i][0] || '').trim();
    if (!id || id.startsWith('⚠️')) continue;
    totalValidos++;
    const mesAno = String(dados[i][3] || '').trim();
    if (!mesAno) { semMesAno++; continue; }
    if (!dentroDoEscopo(mesAno)) { foraEscopo++; continue; }
    const temComentario = String(dados[i][13] || '').trim() === 'TRUE';
    if (temComentario) comComentario++; else semComentario++;
  }

  Logger.log('\n--- RESUMO ---');
  Logger.log('Registros válidos:       ' + totalValidos);
  Logger.log('Sem mes_ano (problema):  ' + semMesAno);
  Logger.log('Fora do escopo (<04/26): ' + foraEscopo);
  Logger.log('Dentro do escopo:        ' + (comComentario + semComentario));
  Logger.log('  → Com comentário (IA): ' + comComentario);
  Logger.log('  → Sem comentário:      ' + semComentario);

  if (semMesAno > 0) {
    Logger.log('\n⚠️  PROBLEMA DETECTADO: ' + semMesAno + ' registros com mes_ano vazio.');
    Logger.log('   → Execute corrigirMesAno() para corrigir automaticamente.');
  } else if (comComentario + semComentario === 0) {
    Logger.log('\n⚠️  Todos os registros estão fora do escopo de datas.');
    Logger.log('   → Verifique o campo DATA_INICIO_CLASSIFICACAO na aba Config.');
    Logger.log('   → Valor atual deveria ser: 01/04/2026');
  } else {
    Logger.log('\n✅ Base parece OK. Execute executarClassificacaoIA() novamente.');
  }

  Logger.log('========================================');
}

// =============================================================================
// CORREÇÃO: recalcula mes_ano e data_resposta para linhas com campo vazio
// =============================================================================
function corrigirMesAno() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaConsolidada = ss.getSheetByName('Base_Consolidada');
  const dados = abaConsolidada.getDataRange().getValues();

  Logger.log('=== CORRIGINDO mes_ano ===');
  let corrigidos = 0;
  let semData = 0;

  for (let i = 1; i < dados.length; i++) {
    const id = String(dados[i][0] || '').trim();
    if (!id || id.startsWith('⚠️')) continue;

    const mesAnoAtual = String(dados[i][3] || '').trim();
    if (mesAnoAtual) continue; // já tem, não precisa corrigir

    // Tenta extrair data da coluna hora (índice 5) ou data_resposta (índice 2)
    let dataExtraida = tentarParsearData(dados[i][5]) || tentarParsearData(dados[i][2]);

    if (dataExtraida) {
      const mesAnoNovo = Utilities.formatDate(dataExtraida, 'America/Fortaleza', 'MM/yyyy');
      const dataFormatada = Utilities.formatDate(dataExtraida, 'America/Fortaleza', 'dd/MM/yyyy');
      const dia = dataExtraida.getDate();

      // Grava mes_ano (col 4 = índice 3), data_resposta (col 3 = índice 2), dia (col 5 = índice 4)
      abaConsolidada.getRange(i + 1, 3).setValue(dataFormatada);
      abaConsolidada.getRange(i + 1, 4).setValue(mesAnoNovo);
      abaConsolidada.getRange(i + 1, 5).setValue(dia);
      corrigidos++;
    } else {
      semData++;
      Logger.log('Linha ' + (i+1) + ': não foi possível extrair data. hora=' + dados[i][5] + ' | data_resposta=' + dados[i][2]);
    }
  }

  Logger.log('Corrigidos: ' + corrigidos);
  Logger.log('Sem data válida: ' + semData);
  Logger.log('=== CORREÇÃO CONCLUÍDA ===');

  if (corrigidos > 0) {
    SpreadsheetApp.getUi().alert('✅ ' + corrigidos + ' registros corrigidos!\n\nAgora execute a Classificação IA novamente.');
  }
}

/**
 * Tenta parsear uma data de vários formatos possíveis.
 * Retorna Date ou null.
 */
function tentarParsearData(valor) {
  if (!valor) return null;
  if (valor instanceof Date && !isNaN(valor.getTime())) return valor;

  const s = String(valor).trim();
  if (!s) return null;

  // Formato brasileiro: dd/MM/yyyy ou dd/MM/yyyy HH:mm
  const matchBR = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (matchBR) {
    const d = new Date(parseInt(matchBR[3]), parseInt(matchBR[2]) - 1, parseInt(matchBR[1]));
    if (!isNaN(d.getTime())) return d;
  }

  // Formato ISO: yyyy-MM-dd ou yyyy/MM/dd
  const matchISO = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (matchISO) {
    const d = new Date(parseInt(matchISO[1]), parseInt(matchISO[2]) - 1, parseInt(matchISO[3]));
    if (!isNaN(d.getTime())) return d;
  }

  // Tenta o parser nativo como último recurso
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;

  return null;
}

// =============================================================================
// VERIFICAR CONFIGURAÇÕES
// =============================================================================
function verificarConfig() {
  const config = lerConfig();
  Logger.log('=== VERIFICAÇÃO DE CONFIG ===');
  Logger.log('API_PROVIDER:               ' + config.API_PROVIDER);
  Logger.log('API_KEY (primeiros 8 chars):' + (config.API_KEY ? config.API_KEY.substring(0, 8) + '...' : '❌ VAZIA'));
  Logger.log('MODELO:                     ' + config.MODELO);
  Logger.log('DATA_INICIO_CLASSIFICACAO:  ' + config.DATA_INICIO_CLASSIFICACAO);
  Logger.log('LOTE_TAMANHO:               ' + config.LOTE_TAMANHO);
  Logger.log('ID_PLANILHA_SOLVIS:         ' + config.ID_PLANILHA_SOLVIS);
  Logger.log('ID_PLANILHA_WEHELP:         ' + config.ID_PLANILHA_WEHELP);

  if (!config.API_KEY) Logger.log('❌ API_KEY ausente — preencha na aba Config!');
  else Logger.log('✅ Config OK');
}

// =============================================================================
// TESTE RÁPIDO DA API GEMINI (1 comentário)
// =============================================================================
function testarAPIIA() {
  const config = lerConfig();
  if (!config.API_KEY) {
    Logger.log('❌ API_KEY não configurada.');
    return;
  }

  Logger.log('Testando conexão com API IA (' + config.API_PROVIDER + ' / ' + config.MODELO + ')...');
  const promptTeste = `Classifique este comentário de academia e retorne JSON:
{"seq":0,"id":"TESTE-001","nota":3,"unidade":"Aldeota","comentario":"Os banheiros estão sempre sujos, já reclamei várias vezes e não melhorou nada."}

Retorne apenas um array JSON com 1 objeto no formato descrito no sistema.
[{"seq":0,"id":"TESTE-001","comentario_sem_sentido":false,"tema_principal":"Limpeza","tema_secundario":"Banheiro/Vestiário","sentimento":"Muito negativo","criticidade":"Alta","tipo_comentario":"Reclamação grave","resumo":"Banheiros sujos, problema recorrente sem solução","requer_acao":true,"risco_cancelamento":false,"palavra_critica":"sempre sujos","acao_sugerida":"Reforçar checklist de limpeza e aumentar frequência de rondas","prioridade_acao":"Alta","problema_recorrente":true,"ocultar_nome_colaborador":false}]`;

  try {
    const resposta = chamarAPIIA(promptTeste, config);
    Logger.log('✅ API IA respondeu com sucesso!');
    Logger.log('Resposta: ' + String(resposta).substring(0, 300));
  } catch (e) {
    Logger.log('❌ Erro ao chamar API IA: ' + e.message);
  }
}

// =============================================================================
// DIAGNÓSTICO: lista todas as colunas das planilhas fonte com amostra de dados
// Execute para identificar qual coluna tem o comentário real
// =============================================================================
function diagnosticarColunasFonte() {
  const config = lerConfig();

  Logger.log('========================================');
  Logger.log('DIAGNÓSTICO — COLUNAS DAS FONTES');
  Logger.log('========================================');

  // Solvis
  try {
    const ssSolvis = SpreadsheetApp.openById(config.ID_PLANILHA_SOLVIS);
    const abaSolvis = ssSolvis.getSheetByName(config.ABA_SOLVIS) || ssSolvis.getSheets()[0];
    const dados = abaSolvis.getRange(1, 1, 3, abaSolvis.getLastColumn()).getValues();
    const cabecalho = dados[0];
    Logger.log('\n--- SOLVIS: colunas e amostra ---');
    cabecalho.forEach((nome, i) => {
      const v1 = String(dados[1] ? dados[1][i] : '').substring(0, 80);
      const v2 = String(dados[2] ? dados[2][i] : '').substring(0, 80);
      Logger.log('[' + i + '] "' + nome + '" | linha1: ' + v1 + ' | linha2: ' + v2);
    });
  } catch (e) {
    Logger.log('Erro ao abrir Solvis: ' + e.message);
  }

  // WeHelp
  try {
    const ssWeHelp = SpreadsheetApp.openById(config.ID_PLANILHA_WEHELP);
    let abaWeHelp = ssWeHelp.getSheetByName(config.ABA_WEHELP);
    if (!abaWeHelp) {
      for (const alt of ['dados', 'Respostas', 'respostas', 'Sheet1', 'Página1']) {
        abaWeHelp = ssWeHelp.getSheetByName(alt);
        if (abaWeHelp) break;
      }
    }
    if (!abaWeHelp) abaWeHelp = ssWeHelp.getSheets()[0];
    const dados = abaWeHelp.getRange(1, 1, 3, abaWeHelp.getLastColumn()).getValues();
    const cabecalho = dados[0];
    Logger.log('\n--- WEHELP: colunas e amostra ---');
    cabecalho.forEach((nome, i) => {
      const v1 = String(dados[1] ? dados[1][i] : '').substring(0, 80);
      const v2 = String(dados[2] ? dados[2][i] : '').substring(0, 80);
      Logger.log('[' + i + '] "' + nome + '" | linha1: ' + v1 + ' | linha2: ' + v2);
    });
  } catch (e) {
    Logger.log('Erro ao abrir WeHelp: ' + e.message);
  }

  Logger.log('\n========================================');
}
