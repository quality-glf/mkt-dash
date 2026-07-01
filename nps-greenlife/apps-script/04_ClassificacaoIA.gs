// =============================================================================
// NPS GREENLIFE ACADEMIAS — CLASSIFICAÇÃO AUTOMÁTICA POR IA
// Arquivo: 04_ClassificacaoIA.gs
// =============================================================================

// =============================================================================
// PONTO DE ENTRADA
// =============================================================================
function executarClassificacaoIA() {
  const inicio = new Date();
  Logger.log('--- Classificação IA iniciada ---');

  const config = lerConfig();
  if (!config.API_KEY) {
    throw new Error('API_KEY não configurada. Insira sua chave Anthropic na aba Config (campo API_KEY).');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaConsolidada = ss.getSheetByName('Base_Consolidada');
  const abaClassificada = ss.getSheetByName('Base_Classificada_IA');

  // IDs já classificados
  const idsClassificados = new Set();
  const classifData = abaClassificada.getDataRange().getValues();
  for (let i = 1; i < classifData.length; i++) {
    const id = String(classifData[i][0] || '').trim();
    if (id && !id.startsWith('⚠️')) idsClassificados.add(id);
  }

  // Carrega registros ainda não classificados
  const consolidadaData = abaConsolidada.getDataRange().getValues();
  const cabecalho = consolidadaData[0];
  const paraClassificar = [];
  const semComentario = [];

  // Data de início vinda do Config (respeita filtro de teste)
  const dataInicio = lerDataInicioClassificacao();

  for (let i = 1; i < consolidadaData.length; i++) {
    const linha = consolidadaData[i];
    const id = String(linha[0] || '').trim();
    if (!id || id.startsWith('⚠️')) continue;
    if (idsClassificados.has(id)) continue;

    const temComentario = linha[13] === true || String(linha[13]).toUpperCase() === 'TRUE';

    // Filtra pela data de início configurada no Config
    const dataRespostaRaw = linha[2];
    let dataResposta = null;
    if (dataRespostaRaw instanceof Date) {
      dataResposta = dataRespostaRaw;
    } else if (dataRespostaRaw) {
      const p = String(dataRespostaRaw).split('/');
      if (p.length === 3) dataResposta = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
    }
    if (!dataResposta || dataResposta < dataInicio) continue;

    if (temComentario) {
      paraClassificar.push(linha);
    } else {
      semComentario.push(linha);
    }
  }

  Logger.log('Registros com comentário para classificar: ' + paraClassificar.length);
  Logger.log('Registros sem comentário (classificação simples): ' + semComentario.length);

  // Nada a fazer — remove trigger e encerra
  if (paraClassificar.length === 0 && semComentario.length === 0) {
    Logger.log('✅ Todos os registros já classificados! Removendo trigger automático.');
    removerTriggerClassificacao();
    return;
  }

  // Aplica limite por execução (MAX_REGISTROS_POR_EXECUCAO da Config, padrão 9999 = sem limite)
  const maxExec = Number(config.MAX_REGISTROS_DIA) || 9999;
  const paraClassificarLimitado = paraClassificar.slice(0, maxExec);
  const semComentarioLimitado = semComentario.slice(0, Math.max(0, maxExec - paraClassificarLimitado.length));
  if (paraClassificar.length > maxExec) {
    Logger.log('⚠️ Limite de ' + maxExec + ' aplicado. Restam ' + (paraClassificar.length - maxExec) + ' para próximas execuções.');
  }

  // Classifica sem comentário (regra simples, sem API)
  const linhasSemComentario = semComentarioLimitado.map(l => classificarSemComentario(l));

  // Classifica com IA em lotes (com proteção de timeout de 5 minutos)
  const lotes = dividirEmLotes(paraClassificarLimitado, Number(config.LOTE_TAMANHO) || 10);
  const linhasComIA = [];
  let erros = 0;
  const TEMPO_LIMITE_MS = 5 * 60 * 1000; // 5 minutos

  for (let idx = 0; idx < lotes.length; idx++) {
    if (new Date() - inicio > TEMPO_LIMITE_MS) {
      Logger.log('⏱ Tempo limite de 5 min atingido após ' + idx + ' lotes. Próxima execução continua de onde parou.');
      break;
    }
    const lote = lotes[idx];
    Logger.log('Processando lote ' + (idx + 1) + '/' + lotes.length + ' (' + lote.length + ' registros)');
    try {
      const resultado = classificarLoteIA(lote, config);
      linhasComIA.push(...resultado);
    } catch (e) {
      Logger.log('Erro no lote ' + (idx + 1) + ': ' + e.message);
      erros++;
      lote.forEach(l => linhasComIA.push(classificarFallback(l, e.message)));
    }
    if (idx < lotes.length - 1) Utilities.sleep(3000); // 3s entre lotes (Claude não tem limite apertado de TPM)
  }

  // Grava tudo na Base_Classificada_IA
  const todasLinhas = [...linhasSemComentario, ...linhasComIA];
  if (todasLinhas.length > 0) {
    // Remove aviso da linha 2 se ainda existir
    const linhaDois = abaClassificada.getRange(2, 1).getValue();
    const inicioGravacao = String(linhaDois).startsWith('⚠️')
      ? 2
      : abaClassificada.getLastRow() + 1;

    // Se primeira gravação, começa na linha 2
    const linhaFinal = idsClassificados.size === 0 ? 2 : abaClassificada.getLastRow() + 1;
    abaClassificada.getRange(linhaFinal, 1, todasLinhas.length, todasLinhas[0].length)
      .setValues(todasLinhas);

    // Formatação condicional básica (cor por criticidade)
    aplicarCoresCriticidade(abaClassificada, linhaFinal, todasLinhas.length);
  }

  // Marca como classificados na Base_Consolidada
  marcarComoClassificados(abaConsolidada, consolidadaData, idsClassificados, todasLinhas);

  const fim = new Date();
  const seg = Math.round((fim - inicio) / 1000);

  registrarLog(
    'CLASSIFICACAO_IA', 'ambas', 0,
    todasLinhas.length, linhasComIA.length,
    linhasSemComentario.length, erros,
    config.VERSAO_PROMPT, config.MODELO, seg,
    erros > 0 ? 'PARCIAL' : 'OK',
    erros > 0 ? erros + ' lotes com erro' : '',
    'Com IA: ' + linhasComIA.length + ' | Sem comentário: ' + linhasSemComentario.length
  );

  atualizarConfigTimestamp('ULTIMO_PROCESSAMENTO_IA', erros > 0 ? 'PARCIAL' : 'OK');
  Logger.log('Classificação IA concluída em ' + seg + 's. Total: ' + todasLinhas.length);
}

// =============================================================================
// CLASSIFICAÇÃO COM CLAUDE (Anthropic)
// =============================================================================
function classificarLoteIA(lote, config) {
  const prompt = montarPromptLote(lote);
  const resposta = chamarAPIIA(prompt, config);
  return parsearRespostaIA(resposta, lote);
}

function chamarAPIIA(prompt, config) {
  const url = 'https://api.anthropic.com/v1/messages';
  const modelo = config.MODELO || 'claude-haiku-4-5-20251001';

  const payload = {
    model: modelo,
    max_tokens: 4096,
    messages: [{ role: 'user', content: prompt }],
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': config.API_KEY,
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code !== 200) throw new Error('Claude API retornou código ' + code + ': ' + body.substring(0, 300));

  const json = JSON.parse(body);
  if (!json.content || !json.content[0]) throw new Error('Claude: resposta sem content.');
  return json.content[0].text;
}

// =============================================================================
// CONSTRUÇÃO DO PROMPT
// =============================================================================
function montarPromptLote(lote) {
  const comentariosJson = lote.map((linha, i) => ({
    seq: i,
    id: linha[0],
    nota: linha[9],
    unidade: linha[7],
    comentario: linha[11],
    avaliacao_recepcao: linha[14],
    avaliacao_professor: linha[15],
    avaliacao_limpeza: linha[16],
    avaliacao_aulas: linha[17],
    avaliacao_equipamentos: linha[18],
    avaliacao_ambiente: linha[19],
  }));

  return `Você é um sistema especializado em classificar feedbacks de academias de ginástica.
Analise os comentários abaixo e retorne um array JSON com um objeto por comentário.

REGRAS OBRIGATÓRIAS:
- Retorne APENAS JSON válido, sem texto adicional, sem markdown
- Mantenha a mesma ordem e o campo "seq" de cada item
- Use SOMENTE os valores listados em cada campo

TEMAS PERMITIDOS (use apenas estes):
Atendimento Recepção | Atendimento Musculação | Professor/Instrutor | Aula Coletiva |
Grade de Aulas | Equipamentos | Manutenção | Limpeza | Climatização | Lotação |
Estrutura | Banheiro/Vestiário | Catraca/Acesso | Aplicativo/Sistema | Cancelamento |
Financeiro/Cobrança | Estacionamento | Comunicação | Elogio Geral | Experiência Geral |
Comentário sem sentido | Sem comentário | Outro

VALORES VÁLIDOS:
- sentimento: "Muito positivo" | "Positivo" | "Neutro" | "Negativo leve" | "Muito negativo"
- criticidade: "Baixa" | "Média" | "Alta" | "Crítica"
- prioridade_acao: "Baixa" | "Média" | "Alta" | "Urgente"
- tipo_comentario: "Elogio" | "Sugestão" | "Reclamação construtiva" | "Reclamação grave" | "Neutro"

REGRAS DE NEGÓCIO:
1. comentario_sem_sentido = true se o texto não fizer sentido no contexto de academia (ex: "ok", ".", "sim", caracteres aleatórios)
2. risco_cancelamento = true se: nota ≤ 6 E criticidade Alta/Crítica, OU se o comentário mencionar: "cancelar", "vou sair", "não volto", "arrependido", "cobrança indevida", "muito insatisfeito", "piorou muito", "não recomendo", "quero cancelar", "nunca mais", "péssimo"
3. requer_acao = true se: criticidade Média ou superior, OU risco_cancelamento = true
4. ocultar_nome_colaborador = true se o comentário mencionar nome de pessoa em contexto de: assédio, abuso sexual, acusação grave, palavrão direcionado a alguém
5. problema_recorrente = true se o comentário usa linguagem de continuidade: "sempre", "toda semana", "nunca arrumam", "há meses", "de novo", "continua", "ainda não", "já reclamei"
6. resumo deve ter no máximo 15 palavras, em português, descrevendo objetivamente o feedback
7. acao_sugerida deve ser prática e específica para uma academia, no máximo 20 palavras
8. Se a nota for 9 ou 10 e o comentário for elogio, criticidade deve ser "Baixa" e requer_acao = false
9. Considere as avaliações sub-numéricas como contexto de apoio, mas priorize o texto do comentário

CONTEXTO: Empresa Greenlife Academias — rede de academias de ginástica no Ceará e São Paulo.

COMENTÁRIOS PARA CLASSIFICAR:
${JSON.stringify(comentariosJson, null, 2)}

ATENÇÃO: Retorne APENAS o array JSON puro. Sem markdown, sem \`\`\`json, sem texto antes ou depois. Somente o array começando com [ e terminando com ].

RETORNE EXATAMENTE ESTE FORMATO (array com ${lote.length} objetos):
[
  {
    "seq": 0,
    "id": "...",
    "comentario_sem_sentido": false,
    "tema_principal": "Limpeza",
    "tema_secundario": "Banheiro/Vestiário",
    "sentimento": "Negativo leve",
    "criticidade": "Média",
    "tipo_comentario": "Reclamação construtiva",
    "resumo": "Cliente insatisfeito com limpeza dos banheiros",
    "requer_acao": true,
    "risco_cancelamento": false,
    "palavra_critica": "banheiro sujo",
    "acao_sugerida": "Reforçar ronda de limpeza e checklist nos banheiros",
    "prioridade_acao": "Média",
    "problema_recorrente": false,
    "ocultar_nome_colaborador": false
  }
]`;
}

// =============================================================================
// PARSING DA RESPOSTA IA
// =============================================================================
function parsearRespostaIA(respostaTexto, lote) {
  const config = lerConfig();
  const agora = Utilities.formatDate(new Date(), 'America/Fortaleza', 'dd/MM/yyyy HH:mm:ss');

  let classificacoes;
  try {
    classificacoes = JSON.parse(respostaTexto);
    if (!Array.isArray(classificacoes)) throw new Error('Resposta não é um array');
  } catch (e) {
    // Tenta extrair JSON de dentro de bloco markdown ```json ... ``` ou texto livre
    const matchMarkdown = respostaTexto.match(/```(?:json)?\s*(\[[\s\S]*?\])\s*```/);
    const matchRaw = respostaTexto.match(/\[[\s\S]*\]/);
    const jsonStr = matchMarkdown ? matchMarkdown[1] : (matchRaw ? matchRaw[0] : null);
    if (jsonStr) {
      try {
        classificacoes = JSON.parse(jsonStr);
      } catch (e2) {
        throw new Error('Não foi possível parsear resposta da IA: ' + e2.message + ' | Trecho: ' + jsonStr.substring(0, 200));
      }
    } else {
      throw new Error('Formato de resposta inválido da IA. Resposta recebida: ' + respostaTexto.substring(0, 300));
    }
  }

  return lote.map((linhaConsolidada, i) => {
    const c = classificacoes.find(x => x.seq === i) || classificarFallbackObj(linhaConsolidada, 'seq não encontrado');

    return montarLinhaClassificada(linhaConsolidada, c, agora, config.VERSAO_PROMPT);
  });
}

function montarLinhaClassificada(linhaConsolidada, c, timestamp, versaoPrompt) {
  return [
    linhaConsolidada[0],  // id_unico
    linhaConsolidada[1],  // plataforma
    linhaConsolidada[2],  // data_resposta
    linhaConsolidada[3],  // mes_ano
    linhaConsolidada[4],  // dia
    linhaConsolidada[6],  // unidade_original
    linhaConsolidada[7],  // unidade_padronizada
    linhaConsolidada[8],  // regiao
    '',                   // clientes_ativos (preenchido por lookup posterior)
    linhaConsolidada[9],  // nota
    linhaConsolidada[10], // tipo_nps
    '',                   // categoria_original (não disponível nas bases atuais)
    linhaConsolidada[11], // comentario_original
    linhaConsolidada[12], // comentario_tratado
    linhaConsolidada[13], // tem_comentario
    c.comentario_sem_sentido ? 'TRUE' : 'FALSE',
    c.tema_principal || 'Outro',
    c.tema_secundario || '',
    c.sentimento || 'Neutro',
    c.criticidade || 'Baixa',
    c.tipo_comentario || 'Neutro',
    c.resumo || '',
    c.requer_acao ? 'TRUE' : 'FALSE',
    c.risco_cancelamento ? 'TRUE' : 'FALSE',
    c.palavra_critica || '',
    c.acao_sugerida || '',
    c.prioridade_acao || 'Baixa',
    c.problema_recorrente ? 'TRUE' : 'FALSE',
    c.ocultar_nome_colaborador ? 'TRUE' : 'FALSE',
    timestamp,
    versaoPrompt,
  ];
}

// =============================================================================
// CLASSIFICAÇÃO SEM COMENTÁRIO (regra simples, sem API)
// =============================================================================
function classificarSemComentario(linhaConsolidada) {
  const nota = Number(linhaConsolidada[9]);
  const agora = Utilities.formatDate(new Date(), 'America/Fortaleza', 'dd/MM/yyyy HH:mm:ss');
  const config = lerConfig();

  let criticidade = 'Baixa';
  let requerAcao = false;
  let riscoCancelamento = false;
  let prioridade = 'Baixa';

  if (nota <= 6) {
    criticidade = nota <= 3 ? 'Alta' : 'Média';
    requerAcao = true;
    riscoCancelamento = nota <= 3;
    prioridade = nota <= 3 ? 'Alta' : 'Média';
  }

  const c = {
    comentario_sem_sentido: false,
    tema_principal: 'Sem comentário',
    tema_secundario: '',
    sentimento: nota >= 9 ? 'Muito positivo' : nota >= 7 ? 'Positivo' : nota >= 5 ? 'Neutro' : 'Muito negativo',
    criticidade,
    tipo_comentario: 'Neutro',
    resumo: nota >= 9 ? 'Promotor sem comentário' : nota >= 7 ? 'Neutro sem comentário' : 'Detrator sem comentário',
    requer_acao: requerAcao,
    risco_cancelamento: riscoCancelamento,
    palavra_critica: '',
    acao_sugerida: requerAcao ? 'Acompanhar evolução do NPS desta unidade' : '',
    prioridade_acao: prioridade,
    problema_recorrente: false,
    ocultar_nome_colaborador: false,
  };

  return montarLinhaClassificada(linhaConsolidada, c, agora, config.VERSAO_PROMPT + '-SIMPLES');
}

// =============================================================================
// FALLBACK (quando IA falha)
// =============================================================================
function classificarFallback(linhaConsolidada, motivo) {
  const agora = Utilities.formatDate(new Date(), 'America/Fortaleza', 'dd/MM/yyyy HH:mm:ss');
  const config = lerConfig();

  const c = classificarFallbackObj(linhaConsolidada, motivo);
  return montarLinhaClassificada(linhaConsolidada, c, agora, config.VERSAO_PROMPT + '-FALLBACK');
}

function classificarFallbackObj(linhaConsolidada, motivo) {
  const nota = Number(linhaConsolidada[9]);
  return {
    comentario_sem_sentido: false,
    tema_principal: 'Outro',
    tema_secundario: '',
    sentimento: nota >= 9 ? 'Muito positivo' : nota >= 7 ? 'Positivo' : 'Neutro',
    criticidade: nota <= 6 ? 'Média' : 'Baixa',
    tipo_comentario: 'Neutro',
    resumo: 'Classificação pendente — erro na IA',
    requer_acao: nota <= 6,
    risco_cancelamento: false,
    palavra_critica: '',
    acao_sugerida: '',
    prioridade_acao: 'Baixa',
    problema_recorrente: false,
    ocultar_nome_colaborador: false,
  };
}

// =============================================================================
// APLICAR CORES POR CRITICIDADE
// =============================================================================
function aplicarCoresCriticidade(aba, linhaInicio, qtd) {
  const cores = { 'Crítica': '#fecaca', 'Alta': '#fed7aa', 'Média': '#fef9c3', 'Baixa': '#f0fdf4' };
  for (let i = 0; i < qtd; i++) {
    const linha = linhaInicio + i;
    const criticidade = String(aba.getRange(linha, 20).getValue()); // coluna criticidade_ia
    const cor = cores[criticidade] || '#ffffff';
    aba.getRange(linha, 1, 1, 31).setBackground(cor);
  }
}

// =============================================================================
// MARCA COMO CLASSIFICADOS NA BASE_CONSOLIDADA
// =============================================================================
function marcarComoClassificados(abaConsolidada, consolidadaData, idsJaClassificados, novasLinhas) {
  const idsNovos = new Set(novasLinhas.map(l => String(l[0])));
  const colClassificado = 25; // coluna _classificado_ia (índice 24, col 25)

  for (let i = 1; i < consolidadaData.length; i++) {
    const id = String(consolidadaData[i][0] || '');
    if (idsNovos.has(id)) {
      abaConsolidada.getRange(i + 1, colClassificado).setValue('TRUE');
    }
  }
}

// =============================================================================
// VERIFICAR ESCOPO DE DATA
// =============================================================================
function dentroDoEscopo(mesAno) {
  if (!mesAno) return false;

  // Objeto Date direto (valor bruto da planilha)
  if (mesAno instanceof Date) {
    if (isNaN(mesAno.getTime())) return false;
    const mes = mesAno.getMonth() + 1;
    const ano = mesAno.getFullYear();
    return ano > 2026 || (ano === 2026 && mes >= 4);
  }

  const s = String(mesAno).trim();
  if (!s) return false;

  // Formato MM/yyyy (ex: 04/2026)
  // Formato dd/MM/yyyy (ex: 01/04/2026) — quando Sheets serializa como string
  const partes = s.split('/');
  if (partes.length === 2) {
    const mes = parseInt(partes[0]);
    const ano = parseInt(partes[1]);
    if (!isNaN(mes) && !isNaN(ano)) return ano > 2026 || (ano === 2026 && mes >= 4);
  } else if (partes.length === 3) {
    const mes = parseInt(partes[1]);
    const ano = parseInt(partes[2]);
    if (!isNaN(mes) && !isNaN(ano)) return ano > 2026 || (ano === 2026 && mes >= 4);
  }

  // Último recurso: string no formato nativo JS Date (ex: "Wed Apr 01 2026 00:00:00 GMT-0300")
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    const mes = d.getMonth() + 1;
    const ano = d.getFullYear();
    return ano > 2026 || (ano === 2026 && mes >= 4);
  }

  return false;
}

// =============================================================================
// AUXILIAR: DIVIDIR EM LOTES
// =============================================================================
function dividirEmLotes(array, tamanho) {
  const lotes = [];
  for (let i = 0; i < array.length; i += tamanho) {
    lotes.push(array.slice(i, i + tamanho));
  }
  return lotes;
}

// =============================================================================
// TRIGGERS — classificação automática a cada 1 minuto
// =============================================================================

/**
 * Ativa trigger de 1 em 1 minuto para classificar em background.
 * Execute esta função UMA VEZ para iniciar o processo.
 * O trigger se remove automaticamente quando tudo estiver classificado.
 */
function ativarTriggerClassificacao() {
  // Remove triggers existentes para evitar duplicatas
  removerTriggerClassificacao();

  ScriptApp.newTrigger('executarClassificacaoIA')
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log('✅ Trigger ativado: executarClassificacaoIA a cada 1 minuto.');
  Logger.log('   Rode removerTriggerClassificacao() para parar manualmente.');
}

/**
 * Remove o trigger de classificação (parada manual ou automática).
 */
function removerTriggerClassificacao() {
  const triggers = ScriptApp.getProjectTriggers();
  let removidos = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'executarClassificacaoIA') {
      ScriptApp.deleteTrigger(t);
      removidos++;
    }
  });
  if (removidos > 0) Logger.log('🗑 ' + removidos + ' trigger(s) de classificação removido(s).');
  else Logger.log('Nenhum trigger de classificação ativo.');
}
