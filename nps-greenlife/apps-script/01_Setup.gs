// =============================================================================
// NPS GREENLIFE ACADEMIAS — FASE 1: SETUP DA PLANILHA CONSOLIDADA
// Arquivo: 01_Setup.gs
// Execução: rode setupCompleto() UMA VEZ para criar todas as abas
// =============================================================================

/**
 * Ponto de entrada principal. Rode esta função uma única vez.
 * Ela cria todas as abas, cabeçalhos e dados iniciais.
 */
function setupCompleto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setName('NPS_Greenlife_Base_Consolidada');

  Logger.log('=== INICIANDO SETUP NPS GREENLIFE ===');

  criarAbaConfig(ss);
  criarAbaRegrasNPS(ss);
  criarAbaDePara(ss);
  criarAbaClientesAtivos(ss);
  criarAbaBaseBrutaSolvis(ss);
  criarAbaBaseBrutaWeHelp(ss);
  criarAbaBaseConsolidada(ss);
  criarAbaBaseClassificadaIA(ss);
  criarAbaLogClassificacao(ss);

  // Remove a aba padrão "Página1" se existir
  removerAbaPadrao(ss);

  // Ordena as abas na sequência correta
  ordenarAbas(ss);

  Logger.log('=== SETUP CONCLUÍDO COM SUCESSO ===');
  SpreadsheetApp.getUi().alert(
    '✅ Setup concluído!\n\n' +
    'Abas criadas:\n' +
    '• Config\n• Regras_NPS\n• De_Para_Unidades\n• Clientes_Ativos\n' +
    '• Base_Bruta_Solvis\n• Base_Bruta_WeHelp\n• Base_Consolidada\n' +
    '• Base_Classificada_IA\n• Log_Classificacao\n\n' +
    'Próximo passo: insira sua chave Gemini na aba Config.'
  );
}

// =============================================================================
// ABA: Config
// =============================================================================
function criarAbaConfig(ss) {
  const aba = obterOuCriarAba(ss, 'Config');
  aba.clearContents();
  aba.setTabColor('#1a73e8');

  const dados = [
    ['Parâmetro', 'Valor', 'Descrição'],
    ['API_PROVIDER', 'gemini', 'Provedor de IA: gemini | claude | openai'],
    ['API_KEY', '', '⚠️ Cole aqui sua chave da API Gemini'],
    ['MODELO', 'gemini-1.5-flash', 'Modelo usado para classificação'],
    ['VERSAO_PROMPT', 'v1.0', 'Versão atual do prompt de classificação'],
    ['LOTE_TAMANHO', '20', 'Quantidade de comentários por chamada à API'],
    ['CLASSIFICAR_APENAS_NOVOS', 'TRUE', 'TRUE = classifica só registros sem classificação'],
    ['DATA_INICIO_CLASSIFICACAO', '01/04/2026', 'Só classifica respostas a partir desta data'],
    ['ID_PLANILHA_SOLVIS', '17dp-H7fkHsWMJ_vF9FZBGhGkYFHfWt_fGx7q_VT9rUI', 'ID da planilha Solvis no Google Drive'],
    ['ID_PLANILHA_WEHELP', '1cXuHJf059RjJsenzqLkDGGV47Trh3u9hxJ_7vQxVCIg', 'ID da planilha WeHelp no Google Drive'],
    ['ABA_SOLVIS', 'Dados', 'Nome da aba de dados na planilha Solvis'],
    ['ABA_WEHELP', 'Dados', 'Nome da aba de dados na planilha WeHelp — confirmar nome real'],
    ['ULTIMO_PROCESSAMENTO_IMPORTACAO', '', 'Preenchido automaticamente pelo script'],
    ['ULTIMO_PROCESSAMENTO_IA', '', 'Preenchido automaticamente pelo script'],
    ['ULTIMO_PROCESSAMENTO_STATUS', '', 'OK | ERRO'],
    ['TRIGGER_ATIVO', 'TRUE', 'TRUE = triggers diários ativos'],
    ['META_NPS', '60', 'Meta oficial de NPS da Greenlife (%)'],
    ['EMAIL_ALERTA', '', 'E-mail para receber alertas de erro (opcional)'],
  ];

  aba.getRange(1, 1, dados.length, dados[0].length).setValues(dados);

  // Formatação
  const header = aba.getRange(1, 1, 1, 3);
  header.setBackground('#00594F').setFontColor('white').setFontWeight('bold');

  aba.getRange(3, 2).setBackground('#fff3cd'); // destaque para campo API_KEY

  aba.setColumnWidth(1, 280);
  aba.setColumnWidth(2, 320);
  aba.setColumnWidth(3, 380);

  // Proteger a aba de edições acidentais (apenas aviso)
  const protecao = aba.protect().setDescription('Config — edite com cuidado');
  protecao.setWarningOnly(true);

  Logger.log('✓ Aba Config criada');
}

// =============================================================================
// ABA: Regras_NPS
// =============================================================================
function criarAbaRegrasNPS(ss) {
  const aba = obterOuCriarAba(ss, 'Regras_NPS');
  aba.clearContents();
  aba.setTabColor('#00594F');

  const dados = [
    ['Parâmetro', 'Valor', 'Descrição'],
    ['META_NPS_GERAL', '60', 'Meta de NPS para toda a rede (%)'],
    ['PROMOTOR_MIN', '9', 'Nota mínima para Promotor'],
    ['PROMOTOR_MAX', '10', 'Nota máxima para Promotor'],
    ['NEUTRO_MIN', '7', 'Nota mínima para Neutro'],
    ['NEUTRO_MAX', '8', 'Nota máxima para Neutro'],
    ['DETRATOR_MIN', '0', 'Nota mínima para Detrator'],
    ['DETRATOR_MAX', '6', 'Nota máxima para Detrator'],
    ['', '', ''],
    ['Parâmetro', 'Valor', 'Regras de Risco de Cancelamento'],
    ['RISCO_NOTA_MAX', '6', 'Nota ≤ este valor já é sinal de risco'],
    ['RISCO_CRITICIDADE', 'Alta,Crítica', 'Criticidades que ativam risco'],
    ['RISCO_TERMOS', 'cancelar,vou sair,não volto,arrependido,cobrança indevida,muito insatisfeito,piorou muito,não recomendo,quero cancelar,vou embora,péssimo,horrível,nunca mais,absurdo', 'Termos detectados no comentário que ativam risco'],
    ['', '', ''],
    ['Parâmetro', 'Valor', 'Regras de Exclusão'],
    ['EXCLUIR_UNIDADES', 'Bootcamp,BOOTCAMP', 'Unidades excluídas de todas as análises'],
    ['EXCLUIR_COMENTARIOS_SENSIVEIS', 'TRUE', 'Ocultar nome de colaborador em comentários sensíveis'],
    ['TERMOS_SENSIVEIS', 'assédio,abuso,sexual,xingamento,ameaça,acusação,palavrão', 'Termos que marcam comentário como sensível'],
    ['', '', ''],
    ['Parâmetro', 'Valor', 'Status vs Meta'],
    ['STATUS_ACIMA_META', '60', 'NPS ≥ este valor = Acima da Meta (verde)'],
    ['STATUS_ATENCAO_MIN', '45', 'NPS entre este e a meta = Atenção (amarelo)'],
    ['STATUS_CRITICO_MAX', '44', 'NPS ≤ este valor = Crítico (vermelho)'],
    ['ALERTA_BAIXO_VOLUME', '5', 'Unidades com menos de N respostas no período = alerta de baixo volume'],
  ];

  aba.getRange(1, 1, dados.length, dados[0].length).setValues(dados);

  // Cabeçalhos em verde
  [1, 10, 15, 20].forEach(row => {
    aba.getRange(row, 1, 1, 3).setBackground('#00594F').setFontColor('white').setFontWeight('bold');
  });

  aba.setColumnWidth(1, 260);
  aba.setColumnWidth(2, 420);
  aba.setColumnWidth(3, 360);

  Logger.log('✓ Aba Regras_NPS criada');
}

// =============================================================================
// ABA: De_Para_Unidades
// =============================================================================
function criarAbaDePara(ss) {
  const aba = obterOuCriarAba(ss, 'De_Para_Unidades');
  aba.clearContents();
  aba.setTabColor('#15803D');

  const cabecalho = [
    ['Texto_Entrada', 'Unidade_Padronizada', 'Regiao', 'Plataforma', 'Ativa', 'Observacao']
  ];

  // Cada linha: [texto original, nome padronizado, região, plataforma (ambas/solvis/wehelp), ativa, obs]
  const mapeamentos = [
    // ---- REGIÃO 1 ----
    ['ALDEOTA',           'Aldeota',          'Região 1', 'ambas',  'S', ''],
    ['Aldeota',           'Aldeota',          'Região 1', 'ambas',  'S', ''],
    ['aldeota',           'Aldeota',          'Região 1', 'ambas',  'S', ''],

    ['GUARARAPES',        'Guararapes',       'Região 1', 'ambas',  'S', ''],
    ['Guararapes',        'Guararapes',       'Região 1', 'ambas',  'S', ''],

    ['NOVA GUARARAPES',   'Nova Guararapes',  'Região 1', 'ambas',  'S', ''],
    ['Nova Guararapes',   'Nova Guararapes',  'Região 1', 'ambas',  'S', ''],
    ['Nova guararapes',   'Nova Guararapes',  'Região 1', 'ambas',  'S', ''],

    ['PONTES VIEIRA',     'Pontes Vieira',    'Região 1', 'ambas',  'S', ''],
    ['Pontes Vieira',     'Pontes Vieira',    'Região 1', 'ambas',  'S', ''],
    ['Pontes vieira',     'Pontes Vieira',    'Região 1', 'ambas',  'S', ''],

    ['RUI BARBOSA',       'Rui Barbosa',      'Região 1', 'ambas',  'S', ''],
    ['Rui Barbosa',       'Rui Barbosa',      'Região 1', 'ambas',  'S', ''],
    ['Rui barbosa',       'Rui Barbosa',      'Região 1', 'ambas',  'S', ''],

    ['RioMar',            'RioMar',           'Região 1', 'solvis', 'S', ''],
    ['Riomar',            'RioMar',           'Região 1', 'wehelp', 'S', ''],
    ['RIOMARR',           'RioMar',           'Região 1', 'ambas',  'S', 'Possível typo'],
    ['Rio Mar',           'RioMar',           'Região 1', 'ambas',  'S', ''],
    ['RIOMAR',            'RioMar',           'Região 1', 'ambas',  'S', ''],

    ['CT',                'CT',               'Região 1', 'solvis', 'S', 'Centro de Treinamento'],
    ['CT Greenlife',      'CT',               'Região 1', 'wehelp', 'S', ''],
    ['Greenlife CT',      'CT',               'Região 1', 'ambas',  'S', ''],
    ['GREENLIFE CT',      'CT',               'Região 1', 'ambas',  'S', ''],
    ['CT GREENLIFE',      'CT',               'Região 1', 'ambas',  'S', ''],
    ['ct',                'CT',               'Região 1', 'ambas',  'S', ''],

    ['Shopping Aldeota',  'Shopping Aldeota', 'Região 1', 'solvis', 'S', ''],
    ['Shopping Ald',      'Shopping Aldeota', 'Região 1', 'wehelp', 'S', 'Abreviado no WeHelp'],
    ['SHOPPING ALDEOTA',  'Shopping Aldeota', 'Região 1', 'ambas',  'S', ''],

    ['PERSONAL',          'Personal',         'Região 1', 'solvis', 'S', ''],
    ['Personal',          'Personal',         'Região 1', 'ambas',  'S', ''],

    // ---- REGIÃO 2 ----
    ['FÁTIMA',            'Fátima',           'Região 2', 'ambas',  'S', ''],
    ['Fátima',            'Fátima',           'Região 2', 'ambas',  'S', ''],
    ['FATIMA',            'Fátima',           'Região 2', 'ambas',  'S', ''],
    ['Fatima',            'Fátima',           'Região 2', 'ambas',  'S', ''],

    ['JOQUEI',            'Jóquei',           'Região 2', 'solvis', 'S', ''],
    ['Jóquei',            'Jóquei',           'Região 2', 'ambas',  'S', ''],
    ['JÓQUEI',            'Jóquei',           'Região 2', 'ambas',  'S', ''],
    ['Joquei',            'Jóquei',           'Região 2', 'ambas',  'S', ''],
    ['joquei',            'Jóquei',           'Região 2', 'ambas',  'S', ''],

    ['KENNEDY',           'Kennedy',          'Região 2', 'ambas',  'S', ''],
    ['Kennedy',           'Kennedy',          'Região 2', 'ambas',  'S', ''],

    ['MARAPONGA',         'Maraponga',        'Região 2', 'ambas',  'S', ''],
    ['Maraponga',         'Maraponga',        'Região 2', 'ambas',  'S', ''],

    ['MONTESE',           'Montese',          'Região 2', 'ambas',  'S', ''],
    ['Montese',           'Montese',          'Região 2', 'ambas',  'S', ''],

    ['PARQUELANDIA',      'Parquelândia',     'Região 2', 'solvis', 'S', ''],
    ['PARQUELÂNDIA',      'Parquelândia',     'Região 2', 'ambas',  'S', ''],
    ['Parquelândia',      'Parquelândia',     'Região 2', 'ambas',  'S', ''],
    ['Parquelandia',      'Parquelândia',     'Região 2', 'ambas',  'S', ''],

    ['PASSARÉ',           'Passaré',          'Região 2', 'ambas',  'S', ''],
    ['Passaré',           'Passaré',          'Região 2', 'ambas',  'S', ''],
    ['PASSARE',           'Passaré',          'Região 2', 'ambas',  'S', ''],
    ['Passare',           'Passaré',          'Região 2', 'ambas',  'S', ''],

    ['CT Norte',          'CT Norte',         'Região 2', 'wehelp', 'S', 'Unidade separada do CT'],
    ['CT NORTE',          'CT Norte',         'Região 2', 'solvis', 'S', ''],
    ['Greenlife CT Norte','CT Norte',         'Região 2', 'wehelp', 'S', ''],
    ['CT NORTE GL',       'CT Norte',         'Região 2', 'ambas',  'S', ''],
    ['ct norte',          'CT Norte',         'Região 2', 'ambas',  'S', ''],

    // ---- REGIÃO 3 ----
    ['CAMBEBA',           'Cambeba',          'Região 3', 'ambas',  'S', ''],
    ['Cambeba',           'Cambeba',          'Região 3', 'ambas',  'S', ''],

    ['EUSÉBIO',           'Eusébio',          'Região 3', 'solvis', 'S', ''],
    ['Eusébio',           'Eusébio',          'Região 3', 'ambas',  'S', ''],
    ['EUSEBIO',           'Eusébio',          'Região 3', 'ambas',  'S', ''],
    ['Eusebio',           'Eusébio',          'Região 3', 'wehelp', 'S', ''],

    ['MARACANAU',         'Maracanaú',        'Região 3', 'solvis', 'S', ''],
    ['Maracanau',         'Maracanaú',        'Região 3', 'ambas',  'S', ''],
    ['MARACANAÚ',         'Maracanaú',        'Região 3', 'ambas',  'S', ''],
    ['Maracanaú',         'Maracanaú',        'Região 3', 'ambas',  'S', ''],

    ['MARANGUAPE',        'Maranguape',       'Região 3', 'ambas',  'S', ''],
    ['Maranguape',        'Maranguape',       'Região 3', 'ambas',  'S', ''],

    ['MESSEJANA',         'Messejana',        'Região 3', 'ambas',  'S', ''],
    ['Messejana',         'Messejana',        'Região 3', 'ambas',  'S', ''],

    ['SUL',               'Sul',              'Região 3', 'ambas',  'S', ''],
    ['Sul',               'Sul',              'Região 3', 'ambas',  'S', ''],
    ['Unidade Sul',       'Sul',              'Região 3', 'ambas',  'S', ''],
    ['UNIDADE SUL',       'Sul',              'Região 3', 'ambas',  'S', ''],

    ['CAUCAIA',           'Caucaia',          'Região 3', 'solvis', 'S', ''],
    ['Caucaia',           'Caucaia',          'Região 3', 'wehelp', 'S', ''],

    // ---- SÃO PAULO ----
    ['TATUAPÉ',           'Tatuapé',          'São Paulo', 'solvis', 'S', ''],
    ['Tatuapé',           'Tatuapé',          'São Paulo', 'wehelp', 'S', ''],
    ['TATUAPE',           'Tatuapé',          'São Paulo', 'ambas',  'S', ''],
    ['Tatuape',           'Tatuapé',          'São Paulo', 'ambas',  'S', ''],

    ['BARRA FUNDA',       'Barra Funda',      'São Paulo', 'solvis', 'S', ''],
    ['Barra Funda',       'Barra Funda',      'São Paulo', 'wehelp', 'S', ''],

    ['MOEMA',             'Moema',            'São Paulo', 'solvis', 'S', ''],
    ['Moema',             'Moema',            'São Paulo', 'wehelp', 'S', ''],

    // ---- EXCLUÍDAS ----
    ['BOOTCAMP',          '[EXCLUIR]',        '',          'ambas',  'N', 'Excluída conforme regras do projeto'],
    ['Bootcamp',          '[EXCLUIR]',        '',          'ambas',  'N', 'Excluída conforme regras do projeto'],
    ['bootcamp',          '[EXCLUIR]',        '',          'ambas',  'N', 'Excluída conforme regras do projeto'],
  ];

  const todasLinhas = cabecalho.concat(mapeamentos);
  aba.getRange(1, 1, todasLinhas.length, 6).setValues(todasLinhas);

  // Cabeçalho
  aba.getRange(1, 1, 1, 6)
    .setBackground('#00594F')
    .setFontColor('white')
    .setFontWeight('bold');

  // Destaque visual por região
  colorirDeParaPorRegiao(aba, mapeamentos);

  // Larguras
  aba.setColumnWidth(1, 220);
  aba.setColumnWidth(2, 200);
  aba.setColumnWidth(3, 100);
  aba.setColumnWidth(4, 100);
  aba.setColumnWidth(5, 60);
  aba.setColumnWidth(6, 280);

  // Congelar cabeçalho
  aba.setFrozenRows(1);

  Logger.log('✓ Aba De_Para_Unidades criada com ' + mapeamentos.length + ' mapeamentos');
}

function colorirDeParaPorRegiao(aba, mapeamentos) {
  const cores = {
    'Região 1': '#dbeafe',
    'Região 2': '#dcfce7',
    'Região 3': '#fef9c3',
    'São Paulo': '#ede9fe',
    '':         '#fee2e2', // excluídas
  };

  mapeamentos.forEach((linha, i) => {
    const regiao = linha[2];
    const cor = cores[regiao] || '#ffffff';
    aba.getRange(i + 2, 1, 1, 6).setBackground(cor);
  });
}

// =============================================================================
// ABA: Clientes_Ativos
// =============================================================================
function criarAbaClientesAtivos(ss) {
  const aba = obterOuCriarAba(ss, 'Clientes_Ativos');
  aba.clearContents();
  aba.setTabColor('#F59E0B');

  const cabecalho = [['Unidade_Padronizada', 'Regiao', 'Competencia', 'Clientes_Ativos', 'Fonte', 'Observacao']];

  // Linha de exemplo para orientar o preenchimento
  const exemplo = [
    ['Aldeota', 'Região 1', '04/2026', '820', 'Manual PPQ', 'Exemplo — substituir pelos dados reais'],
  ];

  const dados = cabecalho.concat(exemplo);
  aba.getRange(1, 1, dados.length, 6).setValues(dados);

  // Cabeçalho
  aba.getRange(1, 1, 1, 6)
    .setBackground('#F59E0B')
    .setFontColor('white')
    .setFontWeight('bold');

  // Linha de exemplo em itálico
  aba.getRange(2, 1, 1, 6).setFontStyle('italic').setFontColor('#94a3b8');

  // Instruções
  aba.getRange(4, 1).setValue('INSTRUÇÕES:');
  aba.getRange(4, 1).setFontWeight('bold').setFontColor('#00594F');
  aba.getRange(5, 1).setValue('• Preencha uma linha por unidade por mês de competência');
  aba.getRange(6, 1).setValue('• Competencia no formato MM/AAAA (ex: 04/2026)');
  aba.getRange(7, 1).setValue('• Clientes_Ativos = total de alunos ativos na unidade naquele mês');
  aba.getRange(8, 1).setValue('• Mantenha os nomes de unidade iguais à coluna Unidade_Padronizada do De_Para_Unidades');
  aba.getRange(9, 1).setValue('• Este dado é atualizado mensalmente pelo time PPQ');

  aba.setColumnWidth(1, 200);
  aba.setColumnWidth(2, 110);
  aba.setColumnWidth(3, 110);
  aba.setColumnWidth(4, 130);
  aba.setColumnWidth(5, 130);
  aba.setColumnWidth(6, 320);

  aba.setFrozenRows(1);

  Logger.log('✓ Aba Clientes_Ativos criada');
}

// =============================================================================
// ABA: Base_Bruta_Solvis
// =============================================================================
function criarAbaBaseBrutaSolvis(ss) {
  const aba = obterOuCriarAba(ss, 'Base_Bruta_Solvis');
  aba.clearContents();
  aba.setTabColor('#64748b');

  const cabecalho = [[
    'Unidade_de_pesquisa',
    'Horario',
    'Nota_NPS',
    'Motivo_Nota',
    'Avaliacao_Recepcao',
    'Avaliacao_Professor',
    'Avaliacao_Limpeza',
    'Avaliacao_Aulas',
    'Avaliacao_Equipamentos',
    'Avaliacao_Ambiente',
    'Comentario',
    'Nome_Cliente',
    'Email_Cliente',
    'Telefone_Cliente',
    'Hora',
    'Regiao_Origem',
    'Mes_Origem',
    '_data_importacao',
    '_id_gerado',
  ]];

  aba.getRange(1, 1, 1, cabecalho[0].length).setValues(cabecalho);
  aba.getRange(1, 1, 1, cabecalho[0].length)
    .setBackground('#475569')
    .setFontColor('white')
    .setFontWeight('bold');

  aba.setFrozenRows(1);

  // Aviso de não editar
  aba.getRange(2, 1).setValue('⚠️ Esta aba é preenchida automaticamente pelo script. Não edite manualmente.');
  aba.getRange(2, 1).setFontColor('#dc2626').setFontStyle('italic');

  aba.setColumnWidth(1, 180);
  aba.setColumnWidth(4, 300);
  aba.setColumnWidth(11, 300);

  Logger.log('✓ Aba Base_Bruta_Solvis criada');
}

// =============================================================================
// ABA: Base_Bruta_WeHelp
// =============================================================================
function criarAbaBaseBrutaWeHelp(ss) {
  const aba = obterOuCriarAba(ss, 'Base_Bruta_WeHelp');
  aba.clearContents();
  aba.setTabColor('#64748b');

  const cabecalho = [[
    'Unidade_de_pesquisa',
    'Horario',
    'Nota_NPS',
    'Nome_Indicacao',
    'Telefone_Indicacao',
    'Avaliacao_Recepcao',
    'Avaliacao_Professor',
    'Avaliacao_Limpeza',
    'Avaliacao_Aulas',
    'Avaliacao_Equipamentos',
    'Avaliacao_Ambiente',
    'Comentario',
    'Nome_Cliente',
    'Email_Cliente',
    'Telefone_Cliente',
    'Hora',
    'Regiao_Origem',
    'Mes_Origem',
    '_data_importacao',
    '_id_gerado',
  ]];

  aba.getRange(1, 1, 1, cabecalho[0].length).setValues(cabecalho);
  aba.getRange(1, 1, 1, cabecalho[0].length)
    .setBackground('#475569')
    .setFontColor('white')
    .setFontWeight('bold');

  aba.setFrozenRows(1);

  aba.getRange(2, 1).setValue('⚠️ Esta aba é preenchida automaticamente pelo script. Não edite manualmente.');
  aba.getRange(2, 1).setFontColor('#dc2626').setFontStyle('italic');

  aba.setColumnWidth(1, 180);
  aba.setColumnWidth(12, 300);

  Logger.log('✓ Aba Base_Bruta_WeHelp criada');
}

// =============================================================================
// ABA: Base_Consolidada
// =============================================================================
function criarAbaBaseConsolidada(ss) {
  const aba = obterOuCriarAba(ss, 'Base_Consolidada');
  aba.clearContents();
  aba.setTabColor('#15803D');

  const cabecalho = [[
    'id_unico',
    'plataforma',
    'data_resposta',
    'mes_ano',
    'dia',
    'hora',
    'unidade_original',
    'unidade_padronizada',
    'regiao',
    'nota',
    'tipo_nps',
    'comentario_original',
    'comentario_tratado',
    'tem_comentario',
    'avaliacao_recepcao',
    'avaliacao_professor',
    'avaliacao_limpeza',
    'avaliacao_aulas',
    'avaliacao_equipamentos',
    'avaliacao_ambiente',
    'nome_cliente',
    'email_cliente',
    'mes_origem',
    '_data_importacao',
    '_classificado_ia',
  ]];

  aba.getRange(1, 1, 1, cabecalho[0].length).setValues(cabecalho);
  aba.getRange(1, 1, 1, cabecalho[0].length)
    .setBackground('#15803D')
    .setFontColor('white')
    .setFontWeight('bold');

  aba.setFrozenRows(1);

  aba.getRange(2, 1).setValue('⚠️ Esta aba é preenchida automaticamente pelo script. Não edite manualmente.');
  aba.getRange(2, 1).setFontColor('#dc2626').setFontStyle('italic');

  aba.setColumnWidth(1, 180);
  aba.setColumnWidth(12, 320);
  aba.setColumnWidth(13, 320);

  Logger.log('✓ Aba Base_Consolidada criada');
}

// =============================================================================
// ABA: Base_Classificada_IA
// =============================================================================
function criarAbaBaseClassificadaIA(ss) {
  const aba = obterOuCriarAba(ss, 'Base_Classificada_IA');
  aba.clearContents();
  aba.setTabColor('#7c3aed');

  const cabecalho = [[
    // Identificação
    'id_unico',
    'plataforma',
    'data_resposta',
    'mes_ano',
    'dia',
    // Unidade
    'unidade_original',
    'unidade_padronizada',
    'regiao',
    'clientes_ativos',
    // Nota e tipo
    'nota',
    'tipo_nps',
    // Categoria e comentário
    'categoria_original',
    'comentario_original',
    'comentario_tratado',
    'tem_comentario',
    'comentario_sem_sentido',
    // Classificação IA — temas e sentimento
    'tema_principal_ia',
    'tema_secundario_ia',
    'sentimento_ia',
    'criticidade_ia',
    'tipo_comentario_ia',
    'resumo_ia',
    // Ações e riscos
    'requer_acao',
    'risco_cancelamento',
    'palavra_critica_detectada',
    'acao_sugerida_ia',
    'prioridade_acao',
    'problema_recorrente',
    'ocultar_nome_colaborador',
    // Controle
    'data_classificacao_ia',
    'versao_prompt_ia',
  ]];

  aba.getRange(1, 1, 1, cabecalho[0].length).setValues(cabecalho);
  aba.getRange(1, 1, 1, cabecalho[0].length)
    .setBackground('#7c3aed')
    .setFontColor('white')
    .setFontWeight('bold');

  // Grupos visuais por cor de coluna (facilita leitura)
  aba.getRange(1, 1, 1, 5).setBackground('#6d28d9');    // identificação
  aba.getRange(1, 6, 1, 4).setBackground('#5b21b6');    // unidade
  aba.getRange(1, 10, 1, 2).setBackground('#4c1d95');   // nota
  aba.getRange(1, 12, 1, 5).setBackground('#7c3aed');   // comentário
  aba.getRange(1, 17, 1, 6).setBackground('#8b5cf6');   // classificação IA
  aba.getRange(1, 23, 1, 7).setBackground('#a78bfa');   // ações
  aba.getRange(1, 30, 1, 2).setBackground('#c4b5fd');   // controle

  aba.setFrozenRows(1);
  aba.setFrozenColumns(2);

  aba.getRange(2, 1).setValue('⚠️ Esta aba é preenchida automaticamente pelo script. Não edite manualmente.');
  aba.getRange(2, 1).setFontColor('#dc2626').setFontStyle('italic');

  // Larguras importantes
  aba.setColumnWidth(1, 180);  // id_unico
  aba.setColumnWidth(13, 350); // comentario_original
  aba.setColumnWidth(14, 300); // comentario_tratado
  aba.setColumnWidth(22, 280); // resumo_ia
  aba.setColumnWidth(26, 320); // acao_sugerida_ia

  Logger.log('✓ Aba Base_Classificada_IA criada (31 colunas)');
}

// =============================================================================
// ABA: Log_Classificacao
// =============================================================================
function criarAbaLogClassificacao(ss) {
  const aba = obterOuCriarAba(ss, 'Log_Classificacao');
  aba.clearContents();
  aba.setTabColor('#94a3b8');

  const cabecalho = [[
    'timestamp',
    'tipo_operacao',
    'plataforma',
    'registros_importados',
    'registros_novos',
    'registros_classificados',
    'registros_sem_comentario',
    'registros_com_erro',
    'versao_prompt',
    'modelo_ia',
    'tempo_execucao_seg',
    'status',
    'mensagem_erro',
    'detalhes',
  ]];

  aba.getRange(1, 1, 1, cabecalho[0].length).setValues(cabecalho);
  aba.getRange(1, 1, 1, cabecalho[0].length)
    .setBackground('#475569')
    .setFontColor('white')
    .setFontWeight('bold');

  aba.setFrozenRows(1);
  aba.setColumnWidth(1, 160);
  aba.setColumnWidth(13, 300);
  aba.setColumnWidth(14, 400);

  Logger.log('✓ Aba Log_Classificacao criada');
}

// =============================================================================
// UTILITÁRIOS
// =============================================================================

function obterOuCriarAba(ss, nome) {
  let aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    Logger.log('  → Aba criada: ' + nome);
  } else {
    Logger.log('  → Aba já existia, limpando: ' + nome);
  }
  return aba;
}

function removerAbaPadrao(ss) {
  const nomesPadrao = ['Página1', 'Sheet1', 'Planilha1', 'Plan1'];
  nomesPadrao.forEach(nome => {
    const aba = ss.getSheetByName(nome);
    if (aba && ss.getSheets().length > 1) {
      ss.deleteSheet(aba);
      Logger.log('  → Aba padrão removida: ' + nome);
    }
  });
}

function ordenarAbas(ss) {
  const ordemDesejada = [
    'Base_Classificada_IA',
    'Base_Consolidada',
    'Base_Bruta_Solvis',
    'Base_Bruta_WeHelp',
    'Clientes_Ativos',
    'De_Para_Unidades',
    'Regras_NPS',
    'Log_Classificacao',
    'Config',
  ];

  ordemDesejada.forEach((nome, posicao) => {
    const aba = ss.getSheetByName(nome);
    if (aba) {
      ss.setActiveSheet(aba);
      ss.moveActiveSheet(posicao + 1);
    }
  });

  Logger.log('✓ Abas ordenadas');
}
