/**
 * Este script realiza uma análise estatística descritiva dos dados
 * da aba de GCR e gera um resumo em uma nova aba de análise.
 * As configurações são lidas do arquivo 'Configuracoes.gs'.
 */

/**
 * Função principal para orquestrar a análise estatística.
 */
move file to src/analise.gs
  try {
    // As constantes são lidas do arquivo Configuracoes.gs
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_PRINCIPAL);
    const abaDados = planilha.getSheetByName(NOME_ABA_DADOS_GCR);

    if (!abaDados) {
      throw new Error(`A aba de dados "${NOME_ABA_DADOS_GCR}" não foi encontrada.`);
    }

    const dadosCompletos = abaDados.getDataRange().getValues();
    if (dadosCompletos.length < 2) {
      Logger.log('Não há dados suficientes para análise.');
      return;
    }

    const cabecalhos = dadosCompletos.shift();
    const indices = mapearIndices(cabecalhos);
    
    // --- Realiza as contagens ---
    const contagemPorGatilho = contarFrequencia(dadosCompletos, indices['Gatilho']);
    const contagemPorCausal = contarFrequencia(dadosCompletos, indices['Causal']);
    const contagemPorClassificacao = contarFrequencia(dadosCompletos, indices['Classificação']);
    const contagemPorRegional = contarFrequencia(dadosCompletos, indices['Regional']);
    const contagemPorMes = contarFrequencia(dadosCompletos, indices['Mês']);
    const contagemStatusPep = contarFrequencia(dadosCompletos, indices['Status PEP']);

    // --- Prepara a aba de resultados ---
    let abaAnalise = planilha.getSheetByName(NOME_ABA_ANALISE_ESTATISTICA);
    if (abaAnalise) {
      abaAnalise.clear();
    } else {
      abaAnalise = planilha.insertSheet(NOME_ABA_ANALISE_ESTATISTICA);
    }

    // --- Escreve os resultados na aba de análise ---
    let linhaAtual = 1;
    linhaAtual = escreverTabela(abaAnalise, linhaAtual, 'Contagem por Gatilho', contagemPorGatilho);
    linhaAtual = escreverTabela(abaAnalise, linhaAtual, 'Contagem por Causal', contagemPorCausal);
    linhaAtual = escreverTabela(abaAnalise, linhaAtual, 'Contagem por Classificação', contagemPorClassificacao);
    linhaAtual = escreverTabela(abaAnalise, linhaAtual, 'Contagem por Regional', contagemPorRegional);
    linhaAtual = escreverTabela(abaAnalise, linhaAtual, 'Contagem por Mês', contagemPorMes);
    linhaAtual = escreverTabela(abaAnalise, linhaAtual, 'Contagem por Status PEP', contagemStatusPep);

    formatarAbaAnalise(abaAnalise);

    SpreadsheetApp.getUi().alert('Sucesso!', `A análise estatística foi concluída e os resultados estão na aba "${NOME_ABA_ANALISE_ESTATISTICA}".`, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Análise estatística concluída com sucesso.');

  } catch (e) {
    Logger.log(`Ocorreu um erro na análise: ${e.message}\nStack: ${e.stack}`);
    SpreadsheetApp.getUi().alert('Erro', `Ocorreu um erro: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Mapeia os nomes das colunas para seus respectivos índices.
 */
function mapearIndices(cabecalhos) {
  const indices = {};
  cabecalhos.forEach((col, i) => {
    indices[col] = i;
  });
  return indices;
}

/**
 * Conta a frequência de valores em uma determinada coluna.
 */
function contarFrequencia(dados, indiceColuna) {
  if (indiceColuna === undefined) return [];
  const contagens = {};
  dados.forEach(linha => {
    const item = linha[indiceColuna] || 'Vazio';
    contagens[item] = (contagens[item] || 0) + 1;
  });
  return Object.entries(contagens).sort((a, b) => b[1] - a[1]);
}

/**
 * Escreve uma tabela de resultados na aba de análise.
 */
function escreverTabela(aba, linhaInicio, titulo, dadosTabela) {
  if (dadosTabela.length === 0) return linhaInicio;

  aba.getRange(linhaInicio, 1).setValue(titulo).setFontWeight('bold').setFontSize(12);
  aba.getRange(linhaInicio + 1, 1, 1, 2).setValues([['Item', 'Contagem']]).setFontWeight('bold');
  aba.getRange(linhaInicio + 2, 1, dadosTabela.length, 2).setValues(dadosTabela);

  return linhaInicio + dadosTabela.length + 3;
}

/**
 * Aplica formatação geral na aba de análise.
 */
function formatarAbaAnalise(aba) {
    aba.autoResizeColumns(1, 2);
    aba.setFrozenRows(1);
}
