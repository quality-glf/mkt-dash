// =============================================================================
// NPS GREENLIFE ACADEMIAS — FASE 1: IMPORTAÇÃO E CONSOLIDAÇÃO
// Arquivo: 02_Importacao.gs
// =============================================================================

// =============================================================================
// PONTO DE ENTRADA — chame esta função pelo trigger ou manualmente
// =============================================================================
function executarImportacaoCompleta() {
  const inicio = new Date();
  Logger.log('=== IMPORTAÇÃO INICIADA: ' + inicio.toISOString() + ' ===');

  try {
    importarSolvis();
    importarWeHelp();
    consolidarBases();
    atualizarClientesAtivos();
    atualizarConfigTimestamp('ULTIMO_PROCESSAMENTO_IMPORTACAO', 'OK');
    Logger.log('=== IMPORTAÇÃO CONCLUÍDA ===');
  } catch (e) {
    atualizarConfigTimestamp('ULTIMO_PROCESSAMENTO_IMPORTACAO', 'ERRO: ' + e.message);
    registrarLog('IMPORTACAO', '', 0, 0, 0, 0, 1, '', '', 0, 'ERRO', e.message, e.stack);
    enviarAlertaErro('Erro na importação: ' + e.message);
    throw e;
  }
}

// =============================================================================
// IMPORTAÇÃO SOLVIS
// =============================================================================
function importarSolvis() {
  Logger.log('--- Importando Solvis ---');
  const config = lerConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDestino = ss.getSheetByName('Base_Bruta_Solvis');
  const dataImportacao = new Date().toISOString();

  // Abre a planilha Solvis
  let ssSolvis;
  try {
    ssSolvis = SpreadsheetApp.openById(config.ID_PLANILHA_SOLVIS);
  } catch (e) {
    throw new Error('Não foi possível abrir a planilha Solvis. Verifique o ID_PLANILHA_SOLVIS na aba Config. Erro: ' + e.message);
  }

  const abaSolvis = ssSolvis.getSheetByName(config.ABA_SOLVIS);
  if (!abaSolvis) {
    throw new Error('Aba "' + config.ABA_SOLVIS + '" não encontrada na planilha Solvis.');
  }

  const dadosBrutos = abaSolvis.getDataRange().getValues();
  if (dadosBrutos.length < 2) {
    Logger.log('Solvis: nenhum dado encontrado.');
    return;
  }

  const cabecalhoOriginal = dadosBrutos[0];
  Logger.log('Solvis - cabeçalho detectado: ' + cabecalhoOriginal.join(' | '));

  // Mapeia colunas pelo nome (busca parcial, case-insensitive)
  const idx = mapearColunasSolvis(cabecalhoOriginal);
  Logger.log('Solvis - índices mapeados: ' + JSON.stringify(idx));

  // Limpa destino (mantém cabeçalho na linha 1)
  const ultimaLinhaDestino = abaDestino.getLastRow();
  if (ultimaLinhaDestino > 1) {
    abaDestino.getRange(2, 1, ultimaLinhaDestino - 1, abaDestino.getLastColumn()).clearContent();
  }

  // Transforma e grava
  const linhasParaGravar = [];
  for (let i = 1; i < dadosBrutos.length; i++) {
    const linha = dadosBrutos[i];
    const nota = linha[idx.nota];
    const unidade = linha[idx.unidade];

    // Ignora linhas completamente vazias
    if (!unidade && (nota === '' || nota === null || nota === undefined)) continue;

    // Ignora Bootcamp
    if (String(unidade).toUpperCase().includes('BOOTCAMP')) continue;

    const id = gerarIdUnico('SOL', i, linha[idx.data] || linha[idx.hora]);

    linhasParaGravar.push([
      String(unidade || '').trim(),                     // Unidade_de_pesquisa
      formatarData(linha[idx.hora] || linha[idx.data]), // Horario
      normalizarNota(nota),                             // Nota_NPS
      String(linha[idx.motivo] || '').trim(),           // Motivo_Nota (comentário principal)
      String(linha[idx.recepcao] || '').trim(),         // Avaliacao_Recepcao
      String(linha[idx.professor] || '').trim(),        // Avaliacao_Professor
      String(linha[idx.limpeza] || '').trim(),          // Avaliacao_Limpeza
      String(linha[idx.aulas] || '').trim(),            // Avaliacao_Aulas
      String(linha[idx.equipamentos] || '').trim(),     // Avaliacao_Equipamentos
      String(linha[idx.ambiente] || '').trim(),         // Avaliacao_Ambiente
      String(linha[idx.comentario] || '').trim(),       // Comentario (campo contato — guardar)
      String(linha[idx.nome] || '').trim(),             // Nome_Cliente
      String(linha[idx.email] || '').trim(),            // Email_Cliente
      String(linha[idx.telefone] || '').trim(),         // Telefone_Cliente
      String(linha[idx.hora] || '').trim(),             // Hora
      String(linha[idx.regiao] || '').trim(),           // Regiao_Origem
      String(linha[idx.mes] || '').trim(),              // Mes_Origem
      dataImportacao,                                   // _data_importacao
      id,                                               // _id_gerado
    ]);
  }

  if (linhasParaGravar.length > 0) {
    abaDestino.getRange(2, 1, linhasParaGravar.length, linhasParaGravar[0].length)
      .setValues(linhasParaGravar);
  }

  Logger.log('Solvis: ' + linhasParaGravar.length + ' registros importados.');
  registrarLog('IMPORTACAO_SOLVIS', 'Solvis', dadosBrutos.length - 1, linhasParaGravar.length, 0, 0, 0, '', '', 0, 'OK', '', '');
}

/**
 * Mapeia índices das colunas Solvis por correspondência parcial de nome.
 * Robusto a pequenas variações de formatação no cabeçalho.
 */
function mapearColunasSolvis(cabecalho) {
  const buscar = (termos) => {
    for (let i = 0; i < cabecalho.length; i++) {
      const col = String(cabecalho[i]).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
      for (const t of termos) {
        if (col.includes(t)) return i;
      }
    }
    return -1;
  };

  return {
    unidade:      buscar(['unidade de pesquisa', 'unidade_de_pesquisa', 'unidade']),
    hora:         buscar(['horario', 'horário', 'hora']),
    data:         buscar(['data']),
    nota:         buscar(['probabilidade', 'nota', 'score', 'nps']),
    motivo:       buscar(['motivo da sua nota', 'motivo', 'qual o motivo']),
    recepcao:     buscar(['recepcao', 'recepção', 'recepcao']),
    professor:    buscar(['atendimento do professor', 'professor']),
    limpeza:      buscar(['limpeza', 'organizacao', 'organização']),
    aulas:        buscar(['aulas coletivas', 'aulas']),
    equipamentos: buscar(['equipamentos']),
    ambiente:     buscar(['ambiente']),
    comentario:   buscar(['comentarios', 'comentário', 'comentario', '10.']),
    nome:         buscar(['nome:']),
    email:        buscar(['e-mail', 'email']),
    telefone:     buscar(['telefone']),
    regiao:       buscar(['regiao', 'região']),
    mes:          buscar(['mes', 'mês']),
    resumo:       buscar(['resumo']),
  };
}

// =============================================================================
// IMPORTAÇÃO WEHELP
// =============================================================================
function importarWeHelp() {
  Logger.log('--- Importando WeHelp ---');
  const config = lerConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDestino = ss.getSheetByName('Base_Bruta_WeHelp');
  const dataImportacao = new Date().toISOString();

  let ssWeHelp;
  try {
    ssWeHelp = SpreadsheetApp.openById(config.ID_PLANILHA_WEHELP);
  } catch (e) {
    throw new Error('Não foi possível abrir a planilha WeHelp. Verifique o ID_PLANILHA_WEHELP na aba Config. Erro: ' + e.message);
  }

  // Tenta pelo nome primeiro; se não encontrar, usa a primeira aba
  let abaWeHelp = ssWeHelp.getSheetByName(config.ABA_WEHELP);
  if (!abaWeHelp) {
    // Tenta nomes alternativos comuns
    const alternativas = ['dados', 'Respostas', 'respostas', 'Sheet1', 'Página1'];
    for (const alt of alternativas) {
      abaWeHelp = ssWeHelp.getSheetByName(alt);
      if (abaWeHelp) {
        Logger.log('WeHelp: aba "' + config.ABA_WEHELP + '" não encontrada. Usando "' + alt + '".');
        break;
      }
    }
  }
  if (!abaWeHelp) {
    throw new Error('Nenhuma aba de dados encontrada na planilha WeHelp. Verifique ABA_WEHELP na Config.');
  }

  const dadosBrutos = abaWeHelp.getDataRange().getValues();
  if (dadosBrutos.length < 2) {
    Logger.log('WeHelp: nenhum dado encontrado.');
    return;
  }

  const cabecalhoOriginal = dadosBrutos[0];
  Logger.log('WeHelp - cabeçalho detectado: ' + cabecalhoOriginal.join(' | '));

  const idx = mapearColunasWeHelp(cabecalhoOriginal);
  Logger.log('WeHelp - índices mapeados: ' + JSON.stringify(idx));

  // Limpa destino
  const ultimaLinhaDestino = abaWeHelp.getLastRow ? abaDestino.getLastRow() : 1;
  if (ultimaLinhaDestino > 1) {
    abaDestino.getRange(2, 1, ultimaLinhaDestino - 1, abaDestino.getLastColumn()).clearContent();
  }

  const linhasParaGravar = [];
  for (let i = 1; i < dadosBrutos.length; i++) {
    const linha = dadosBrutos[i];
    const nota = linha[idx.nota];
    const unidade = linha[idx.unidade];

    if (!unidade && (nota === '' || nota === null || nota === undefined)) continue;
    if (String(unidade).toUpperCase().includes('BOOTCAMP')) continue;

    const id = gerarIdUnico('WH', i, linha[idx.hora] || '');

    linhasParaGravar.push([
      String(unidade || '').trim(),
      formatarData(linha[idx.hora] || ''),
      normalizarNota(nota),
      String(linha[idx.nome_indicacao] || '').trim(),
      String(linha[idx.telefone_indicacao] || '').trim(),
      String(linha[idx.recepcao] || '').trim(),
      String(linha[idx.professor] || '').trim(),
      String(linha[idx.limpeza] || '').trim(),
      String(linha[idx.aulas] || '').trim(),
      String(linha[idx.equipamentos] || '').trim(),
      String(linha[idx.ambiente] || '').trim(),
      String(linha[idx.comentario] || '').trim(),   // Comentario (campo 10 — comentário principal no WeHelp)
      String(linha[idx.nome] || '').trim(),
      String(linha[idx.email] || '').trim(),
      String(linha[idx.telefone] || '').trim(),
      String(linha[idx.hora] || '').trim(),
      String(linha[idx.regiao] || '').trim(),
      String(linha[idx.mes] || '').trim(),
      dataImportacao,
      id,
    ]);
  }

  if (linhasParaGravar.length > 0) {
    abaDestino.getRange(2, 1, linhasParaGravar.length, linhasParaGravar[0].length)
      .setValues(linhasParaGravar);
  }

  Logger.log('WeHelp: ' + linhasParaGravar.length + ' registros importados.');
  registrarLog('IMPORTACAO_WEHELP', 'WeHelp', dadosBrutos.length - 1, linhasParaGravar.length, 0, 0, 0, '', '', 0, 'OK', '', '');
}

function mapearColunasWeHelp(cabecalho) {
  const buscar = (termos) => {
    for (let i = 0; i < cabecalho.length; i++) {
      const col = String(cabecalho[i]).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
      for (const t of termos) {
        if (col.includes(t)) return i;
      }
    }
    return -1;
  };

  return {
    unidade:           buscar(['unidade de pesquisa', 'unidade']),
    hora:              buscar(['horario', 'horário', 'hora', 'data']),
    nota:              buscar(['probabilidade', 'nota', 'score', 'nps']),
    nome_indicacao:    buscar(['gostaria de indicar', 'indicar um amigo', 'nome do amigo']),
    telefone_indicacao:buscar(['telefone do amigo', '3. telefone']),
    recepcao:          buscar(['recepcao', 'recepção']),
    professor:         buscar(['atendimento do professor', 'professor']),
    limpeza:           buscar(['limpeza', 'organizacao']),
    aulas:             buscar(['aulas coletivas', 'aulas']),
    equipamentos:      buscar(['equipamentos']),
    ambiente:          buscar(['ambiente']),
    comentario:        buscar(['comentarios', 'comentario', '10. coment']),
    nome:              buscar(['11. nome', 'nome:']),
    email:             buscar(['e-mail', 'email', '12.']),
    telefone:          buscar(['13. telefone', 'telefone*']),
    regiao:            buscar(['regiao', 'região']),
    mes:               buscar(['mes', 'mês']),
  };
}

// =============================================================================
// CONSOLIDAÇÃO DAS DUAS BASES
// =============================================================================
function consolidarBases() {
  Logger.log('--- Consolidando bases ---');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaConsolidada = ss.getSheetByName('Base_Consolidada');
  const abaDePara = ss.getSheetByName('De_Para_Unidades');
  const abaClientesAtivos = ss.getSheetByName('Clientes_Ativos');

  // Carrega De_Para em memória
  const deParaData = abaDePara.getDataRange().getValues();
  const mapaDePara = {};
  for (let i = 1; i < deParaData.length; i++) {
    const entrada = String(deParaData[i][0]).trim();
    const padronizado = String(deParaData[i][1]).trim();
    const regiao = String(deParaData[i][2]).trim();
    if (entrada) {
      mapaDePara[entrada.toLowerCase()] = { unidade: padronizado, regiao: regiao };
    }
  }

  // Carrega clientes ativos em memória
  const clientesData = abaClientesAtivos.getDataRange().getValues();
  const mapaClientes = {};
  for (let i = 1; i < clientesData.length; i++) {
    const unid = String(clientesData[i][0]).trim();
    const competencia = String(clientesData[i][2]).trim();
    const qtd = clientesData[i][3];
    if (unid && competencia) {
      mapaClientes[unid + '|' + competencia] = qtd;
    }
  }

  // Carrega IDs e chaves de conteúdo já existentes na consolidada para não duplicar.
  // A chave de conteúdo (unidade+data+nota) é a segunda barreira contra duplicatas
  // no caso de IDs terem sido gerados de forma não-determinística em versões anteriores.
  const idsExistentes = new Set();
  const conteudoExistente = new Set();
  const consolidadaData = abaConsolidada.getDataRange().getValues();
  for (let i = 1; i < consolidadaData.length; i++) {
    const id = consolidadaData[i][0];
    if (id && !String(id).startsWith('⚠️')) {
      idsExistentes.add(String(id));
      // chave de conteúdo: unidade_padronizada(col8) + data(col3) + nota(col10)
      // Normaliza a data pois o Sheets pode retorná-la como Date object em vez de string
      let dataExistente = consolidadaData[i][2];
      if (dataExistente instanceof Date) {
        dataExistente = Utilities.formatDate(dataExistente, 'America/Fortaleza', 'dd/MM/yyyy');
      } else {
        dataExistente = String(dataExistente || '');
      }
      const chaveConteudo = String(consolidadaData[i][7] || '') + '|' +
                            dataExistente + '|' +
                            String(consolidadaData[i][9] || '');
      if (chaveConteudo !== '||') conteudoExistente.add(chaveConteudo);
    }
  }

  const linhasNovas = [];
  const dataInicio = lerDataInicioClassificacao();

  // Processa Solvis
  const solvisData = ss.getSheetByName('Base_Bruta_Solvis').getDataRange().getValues();
  for (let i = 1; i < solvisData.length; i++) {
    const linha = solvisData[i];
    const id = String(linha[18] || '').trim();                // _id_gerado (coluna 19)
    if (!id || String(linha[0]).startsWith('⚠️')) continue;
    if (idsExistentes.has(id)) continue;

    const dataResposta = extrairData(linha[1]);               // Horario
    if (dataResposta && dataResposta < dataInicio) continue;  // Antes do escopo

    const unidadeOriginal = String(linha[0] || '').trim();
    const { unidadePad, regiao } = resolverDePara(unidadeOriginal, mapaDePara);
    if (unidadePad === '[EXCLUIR]') continue;

    // Segunda barreira: dedup por conteúdo (cobre IDs gerados com timestamp antigo)
    const dataStr = dataResposta ? Utilities.formatDate(dataResposta, 'America/Fortaleza', 'dd/MM/yyyy') : String(linha[1] || '');
    const chaveConteudo = unidadePad + '|' + dataStr + '|' + String(linha[2] || '');
    if (conteudoExistente.has(chaveConteudo)) continue;
    conteudoExistente.add(chaveConteudo);

    const nota = Number(linha[2]);
    const mesAno = formatarMesAno(dataResposta);
    const clientesAtivos = mapaClientes[unidadePad + '|' + mesAno] || '';

    // Comentário principal do Solvis está na coluna K (índice 10) da Base_Bruta_Solvis
    const comentario = String(linha[10] || '').trim();

    linhasNovas.push(montarLinhaConsolidada(
      id, 'Solvis', dataResposta, mesAno, linha[1],
      unidadeOriginal, unidadePad, regiao, nota, comentario,
      linha[4], linha[5], linha[6], linha[7], linha[8], linha[9],
      linha[11], linha[12]
    ));
  }

  // Processa WeHelp
  const webhelpData = ss.getSheetByName('Base_Bruta_WeHelp').getDataRange().getValues();
  for (let i = 1; i < webhelpData.length; i++) {
    const linha = webhelpData[i];
    const id = String(linha[19] || '').trim();                // _id_gerado (coluna 20)
    if (!id || String(linha[0]).startsWith('⚠️')) continue;
    if (idsExistentes.has(id)) continue;

    const dataResposta = extrairData(linha[1]);
    if (dataResposta && dataResposta < dataInicio) continue;

    const unidadeOriginal = String(linha[0] || '').trim();
    const { unidadePad, regiao } = resolverDePara(unidadeOriginal, mapaDePara);
    if (unidadePad === '[EXCLUIR]') continue;

    // Segunda barreira: dedup por conteúdo
    const dataStr = dataResposta ? Utilities.formatDate(dataResposta, 'America/Fortaleza', 'dd/MM/yyyy') : String(linha[1] || '');
    const chaveConteudo = unidadePad + '|' + dataStr + '|' + String(linha[2] || '');
    if (conteudoExistente.has(chaveConteudo)) continue;
    conteudoExistente.add(chaveConteudo);

    const nota = Number(linha[2]);
    const mesAno = formatarMesAno(dataResposta);

    // Comentário principal do WeHelp é o campo "10. Comentarios" (índice 11)
    const comentario = String(linha[11] || '').trim();

    linhasNovas.push(montarLinhaConsolidada(
      id, 'WeHelp', dataResposta, mesAno, linha[1],
      unidadeOriginal, unidadePad, regiao, nota, comentario,
      linha[5], linha[6], linha[7], linha[8], linha[9], linha[10],
      linha[12], linha[13]
    ));
  }

  // Grava novas linhas
  if (linhasNovas.length > 0) {
    const primeiraLinhaVazia = abaConsolidada.getLastRow() + 1;
    // Remove linha de aviso se ainda estiver lá
    const linhaAtual = abaConsolidada.getRange(2, 1).getValue();
    const offset = String(linhaAtual).startsWith('⚠️') ? 2 : primeiraLinhaVazia;
    const linhaGravacao = idsExistentes.size === 0 ? 2 : primeiraLinhaVazia;

    abaConsolidada.getRange(linhaGravacao, 1, linhasNovas.length, linhasNovas[0].length)
      .setValues(linhasNovas);

    Logger.log('Consolidação: ' + linhasNovas.length + ' novos registros adicionados.');
    registrarLog('CONSOLIDACAO', 'ambas', 0, linhasNovas.length, linhasNovas.length, 0, 0, '', '', 0, 'OK', '', '');
  } else {
    Logger.log('Consolidação: nenhum registro novo.');
  }
}

function montarLinhaConsolidada(
  id, plataforma, dataResposta, mesAno, horaOriginal,
  unidadeOriginal, unidadePad, regiao, nota, comentario,
  avalRecepcao, avalProf, avalLimpeza, avalAulas, avalEquip, avalAmb,
  nomeCliente, emailCliente
) {
  const tipoNPS = calcularTipoNPS(nota);
  const comentarioTratado = tratarComentario(comentario);
  const temComentario = comentarioTratado.length > 3;
  const dia = dataResposta ? dataResposta.getDate() : '';

  return [
    id,
    plataforma,
    dataResposta ? Utilities.formatDate(dataResposta, 'America/Fortaleza', 'dd/MM/yyyy') : '',
    mesAno,
    dia,
    String(horaOriginal || '').trim(),
    unidadeOriginal,
    unidadePad,
    regiao,
    nota,
    tipoNPS,
    comentario,
    comentarioTratado,
    temComentario ? 'TRUE' : 'FALSE',
    avalRecepcao,
    avalProf,
    avalLimpeza,
    avalAulas,
    avalEquip,
    avalAmb,
    nomeCliente,
    emailCliente,
    mesAno,
    new Date().toISOString(),
    'FALSE', // _classificado_ia — será marcado TRUE após classificação
  ];
}

// =============================================================================
// ATUALIZAÇÃO DE CLIENTES ATIVOS (cria registros faltantes)
// =============================================================================
function atualizarClientesAtivos() {
  // Esta função não preenche dados (isso é manual pelo PPQ),
  // mas verifica se faltam registros e registra no log.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaConsolidada = ss.getSheetByName('Base_Consolidada');
  const abaClientes = ss.getSheetByName('Clientes_Ativos');

  const consolidadaData = abaConsolidada.getDataRange().getValues();
  const clientesData = abaClientes.getDataRange().getValues();

  const clientesSet = new Set();
  for (let i = 1; i < clientesData.length; i++) {
    clientesSet.add(clientesData[i][0] + '|' + clientesData[i][2]);
  }

  const faltando = new Set();
  for (let i = 1; i < consolidadaData.length; i++) {
    const unid = consolidadaData[i][7]; // unidade_padronizada
    const mesAno = consolidadaData[i][3]; // mes_ano
    if (unid && mesAno && !clientesSet.has(unid + '|' + mesAno)) {
      faltando.add(unid + '|' + mesAno);
    }
  }

  if (faltando.size > 0) {
    Logger.log('⚠️ Clientes_Ativos: ' + faltando.size + ' combinações unidade/mês sem dado de clientes ativos:');
    faltando.forEach(k => Logger.log('  - ' + k));
    registrarLog('AVISO_CLIENTES_ATIVOS', '', 0, 0, 0, 0, 0, '', '', 0,
      'AVISO', 'Preencher Clientes_Ativos para: ' + Array.from(faltando).join('; '), '');
  }
}

// =============================================================================
// FUNÇÕES AUXILIARES
// =============================================================================

function resolverDePara(unidadeOriginal, mapaDePara) {
  const chave = unidadeOriginal.trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');

  // Busca exata primeiro
  if (mapaDePara[chave]) {
    return { unidadePad: mapaDePara[chave].unidade, regiao: mapaDePara[chave].regiao };
  }

  // Busca parcial (para variações não mapeadas)
  for (const [k, v] of Object.entries(mapaDePara)) {
    if (chave.includes(k) || k.includes(chave)) {
      Logger.log('De_Para: correspondência parcial "' + unidadeOriginal + '" → "' + v.unidade + '"');
      return { unidadePad: v.unidade, regiao: v.regiao };
    }
  }

  // Não encontrou
  Logger.log('⚠️ De_Para: unidade sem mapeamento: "' + unidadeOriginal + '"');
  registrarLog('AVISO_DEPARA', '', 0, 0, 0, 0, 0, '', '', 0,
    'AVISO', 'Unidade sem mapeamento: ' + unidadeOriginal, '');
  return { unidadePad: unidadeOriginal, regiao: 'Sem Região' };
}

function calcularTipoNPS(nota) {
  const n = Number(nota);
  if (isNaN(n)) return 'Inválido';
  if (n >= 9) return 'Promotor';
  if (n >= 7) return 'Neutro';
  return 'Detrator';
}

function tratarComentario(texto) {
  if (!texto) return '';
  return String(texto)
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')  // remove acentos
    .replace(/[^\w\s]/g, ' ')          // remove pontuação especial
    .replace(/\s+/g, ' ')              // normaliza espaços
    .trim();
}

function normalizarNota(valor) {
  const n = Number(valor);
  if (isNaN(n) || n < 0 || n > 10) return null;
  return Math.round(n);
}

function gerarIdUnico(prefixo, linha, dataOuHora) {
  // ID determinístico: mesmo registro sempre gera mesmo ID, independente de quando roda.
  // Usa MD5(prefixo + linha + digitos_da_data) para evitar colisão e manter estabilidade.
  const digitos = String(dataOuHora || '').replace(/[^0-9]/g, '').slice(0, 12);
  const chave = prefixo + String(linha) + digitos;
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, chave);
  const hex = bytes.map(function(b) {
    return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0');
  }).join('').slice(0, 6);
  const l = String(linha).padStart(5, '0');
  return prefixo + '-' + l + '-' + hex;
}

function formatarData(valor) {
  if (!valor) return '';
  if (valor instanceof Date) return Utilities.formatDate(valor, 'America/Fortaleza', 'dd/MM/yyyy HH:mm');
  return String(valor);
}

function extrairData(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return valor;
  // Tenta converter string para Date
  const d = new Date(valor);
  if (!isNaN(d.getTime())) return d;
  // Tenta formato brasileiro dd/MM/yyyy
  const match = String(valor).match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (match) return new Date(match[3], match[2] - 1, match[1]);
  return null;
}

function formatarMesAno(data) {
  if (!data) return '';
  if (data instanceof Date) {
    return Utilities.formatDate(data, 'America/Fortaleza', 'MM/yyyy');
  }
  return '';
}

function lerDataInicioClassificacao() {
  const config = lerConfig();
  const raw = config.DATA_INICIO_CLASSIFICACAO;
  if (!raw) return new Date(2026, 3, 1); // fallback: 01/04/2026

  // Formato dd/MM/yyyy (ex: "09/06/2026")
  const partes = raw.split('/');
  if (partes.length === 3) {
    const d = new Date(parseInt(partes[2]), parseInt(partes[1]) - 1, parseInt(partes[0]));
    if (!isNaN(d.getTime())) return d;
  }

  // Sheets converteu a data para string JS (ex: "Mon Jun 09 2026 00:00:00 GMT-0300")
  const d = new Date(raw);
  if (!isNaN(d.getTime())) return d;

  return new Date(2026, 3, 1);
}
