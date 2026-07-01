// =============================================================================
// NPS GREENLIFE ACADEMIAS — UTILITÁRIOS COMPARTILHADOS
// Arquivo: 03_Utils.gs
// =============================================================================

// =============================================================================
// LEITURA DA CONFIG
// =============================================================================

/**
 * Lê todos os parâmetros da aba Config e retorna como objeto.
 * Resultado é cacheado por 1 minuto para evitar leituras repetidas.
 */
let _configCache = null;
let _configCacheTime = 0;

function lerConfig() {
  const agora = new Date().getTime();
  if (_configCache && (agora - _configCacheTime) < 60000) return _configCache;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('Config');
  if (!aba) throw new Error('Aba Config não encontrada. Execute setupCompleto() primeiro.');

  const dados = aba.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < dados.length; i++) {
    const chave = String(dados[i][0] || '').trim();
    const valor = String(dados[i][1] || '').trim();
    if (chave) config[chave] = valor;
  }

  _configCache = config;
  _configCacheTime = agora;
  return config;
}

function atualizarConfigTimestamp(parametro, status) {
  _configCache = null; // invalida cache
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('Config');
  const dados = aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const chave = String(dados[i][0] || '').trim();
    if (chave === parametro) {
      const ts = Utilities.formatDate(new Date(), 'America/Fortaleza', 'dd/MM/yyyy HH:mm:ss');
      aba.getRange(i + 1, 2).setValue(ts);
    }
    if (chave === 'ULTIMO_PROCESSAMENTO_STATUS') {
      aba.getRange(i + 1, 2).setValue(status);
    }
  }
}

// =============================================================================
// LOG DE CLASSIFICAÇÃO
// =============================================================================

function registrarLog(
  tipoOperacao, plataforma,
  registrosImportados, registrosNovos, registrosClassificados,
  registrosSemComentario, registrosComErro,
  versaoPrompt, modeloIA, tempoExecSeg,
  status, mensagemErro, detalhes
) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName('Log_Classificacao');
    if (!aba) return;

    const ts = Utilities.formatDate(new Date(), 'America/Fortaleza', 'dd/MM/yyyy HH:mm:ss');
    const novaLinha = [
      ts, tipoOperacao, plataforma,
      registrosImportados, registrosNovos, registrosClassificados,
      registrosSemComentario, registrosComErro,
      versaoPrompt, modeloIA, tempoExecSeg,
      status, mensagemErro, detalhes
    ];

    const proxLinha = aba.getLastRow() + 1;
    aba.getRange(proxLinha, 1, 1, novaLinha.length).setValues([novaLinha]);

    // Cor por status
    if (status === 'ERRO') {
      aba.getRange(proxLinha, 1, 1, novaLinha.length).setBackground('#fee2e2');
    } else if (status === 'AVISO') {
      aba.getRange(proxLinha, 1, 1, novaLinha.length).setBackground('#fffbeb');
    } else {
      aba.getRange(proxLinha, 1, 1, novaLinha.length).setBackground('#f0fdf4');
    }
  } catch (e) {
    Logger.log('Erro ao registrar log: ' + e.message);
  }
}

// =============================================================================
// ALERTAS POR E-MAIL (opcional)
// =============================================================================

function enviarAlertaErro(mensagem) {
  try {
    const config = lerConfig();
    const email = config.EMAIL_ALERTA;
    if (!email) return;

    MailApp.sendEmail({
      to: email,
      subject: '⚠️ NPS Greenlife — Erro no processamento',
      body: mensagem + '\n\nData: ' + new Date().toString(),
    });
  } catch (e) {
    Logger.log('Falha ao enviar e-mail de alerta: ' + e.message);
  }
}

// =============================================================================
// TRIGGERS
// =============================================================================

/**
 * Configura os dois triggers diários (06h e 18h).
 * Execute esta função uma única vez após o setup.
 */
function configurarTriggers() {
  // Remove triggers antigos para evitar duplicatas
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'executarPipelineCompleto') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Trigger das 06h
  ScriptApp.newTrigger('executarPipelineCompleto')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .inTimezone('America/Fortaleza')
    .create();

  // Trigger das 18h
  ScriptApp.newTrigger('executarPipelineCompleto')
    .timeBased()
    .atHour(18)
    .everyDays(1)
    .inTimezone('America/Fortaleza')
    .create();

  Logger.log('✓ Triggers configurados: 06h e 18h (horário de Fortaleza)');
  SpreadsheetApp.getUi().alert('✅ Triggers configurados!\n\nO pipeline rodará automaticamente às 06h e 18h (horário de Fortaleza).');
}

function removerTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('Todos os triggers removidos.');
}

/**
 * Pipeline completo: importação → consolidação → classificação IA → log
 * Esta é a função chamada pelos triggers automáticos.
 */
function executarPipelineCompleto() {
  const inicio = new Date();
  Logger.log('=== PIPELINE COMPLETO INICIADO: ' + inicio.toISOString() + ' ===');

  try {
    executarImportacaoCompleta();
    executarClassificacaoIA();
    const fim = new Date();
    const seg = Math.round((fim - inicio) / 1000);
    Logger.log('=== PIPELINE CONCLUÍDO em ' + seg + 's ===');
  } catch (e) {
    Logger.log('=== PIPELINE COM ERRO: ' + e.message + ' ===');
    enviarAlertaErro('Erro no pipeline completo: ' + e.message);
  }
}

// =============================================================================
// MENU PERSONALIZADO
// =============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏋️ NPS Greenlife')
    .addItem('▶ Executar pipeline completo', 'executarPipelineCompleto')
    .addSeparator()
    .addItem('↓ Importar apenas Solvis', 'importarSolvis')
    .addItem('↓ Importar apenas WeHelp', 'importarWeHelp')
    .addItem('⚙ Consolidar bases', 'consolidarBases')
    .addSeparator()
    .addItem('🤖 Classificar comentários (IA)', 'executarClassificacaoIA')
    .addSeparator()
    .addItem('⏰ Configurar triggers diários', 'configurarTriggers')
    .addItem('🛑 Remover triggers', 'removerTriggers')
    .addSeparator()
    .addItem('🔧 Setup inicial (só rodar 1x)', 'setupCompleto')
    .addToUi();
}
