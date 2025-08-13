/**
 *
 * Este script finaliza a análise diagnóstica, cruzando os dados
 * de 'Gatilho' com 'Mês' para identificar tendências e sazonalidade.
 * As configurações são lidas do arquivo 'Configuracoes.gs'.
 */
move file to src/analise_temporal.gs
/**
 * Função principal para orquestrar a análise Gatilho vs. Mês.
 */
function executarAnaliseTemporal() {
  try {
    // As constantes são lidas do arquivo Configuracoes.gs
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_PRINCIPAL);
    const abaDados = planilha.getSheetByName(NOME_ABA_DADOS_GCR);

    if (!abaDados) {
      throw new Error(`A aba de dados "${NOME_ABA_DADOS_GCR}" não foi encontrada.`);
    }

    const dadosCompletos = abaDados.getDataRange().getValues();
    if (dadosCompletos.length < 2) {
      Logger.log('Não há dados suficientes para a análise temporal.');
      return;
    }

    const cabecalhos = dadosCompletos.shift();

    const indices = {};
    cabecalhos.forEach((col, i) => { indices[col] = i; });
    
    const indiceGatilho = indices['Gatilho'];
    const indiceMes = indices['Mês'];

    if (indiceGatilho === undefined || indiceMes === undefined) {
        throw new Error("As colunas 'Gatilho' e/ou 'Mês' não foram encontradas.");
    }

    const analiseCruzada = contarAnaliseCruzadaGeneric(dadosCompletos, indiceGatilho, indiceMes);

    let abaResultado = planilha.getSheetByName(NOME_ABA_ANALISE_DIAGNOSTICA);
    if (!abaResultado) {
      throw new Error(`A aba de resultados "${NOME_ABA_ANALISE_DIAGNOSTICA}" não foi encontrada. Por favor, execute as análises anteriores primeiro.`);
    }

    const proximaLinha = abaResultado.getLastRow() + 3;

    escreverTabelaCruzadaGeneric(abaResultado, proximaLinha, 'Análise de Tendência: Gatilho vs. Mês', analiseCruzada, 'Gatilho', 'Mês');

    SpreadsheetApp.getUi().alert('Sucesso!', `A análise de tendência foi concluída e adicionada à aba "${NOME_ABA_ANALISE_DIAGNOSTICA}".`, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Análise de Tendência concluída com sucesso.');

  } catch (e) {
    Logger.log(`Ocorreu um erro na análise temporal: ${e.message}\nStack: ${e.stack}`);
    SpreadsheetApp.getUi().alert('Erro', `Ocorreu um erro: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Conta a frequência de uma coluna de valor dentro de uma coluna de chave.
 * @param {Array<Array<string>>} dados - Os dados da planilha.
 * @param {number} indiceChave - O índice da coluna principal (ex: Gatilho).
 * @param {number} indiceValor - O índice da coluna secundária (ex: Mês).
 * @returns {Object} Um objeto aninhado: { Chave: { Valor: contagem } }.
 */
function contarAnaliseCruzadaGeneric(dados, indiceChave, indiceValor) {
  const contagens = {};
  dados.forEach(linha => {
    const chave = linha[indiceChave] || 'Vazio';
    const valor = linha[indiceValor] || 'Vazio';

    if (!contagens[chave]) {
      contagens[chave] = {};
    }
    contagens[chave][valor] = (contagens[chave][valor] || 0) + 1;
  });
  return contagens;
}

/**
 * Escreve uma tabela de análise cruzada genérica na aba de resultados.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} aba - A aba para escrever.
 * @param {number} linhaInicio - A linha para começar a escrever.
 * @param {string} titulo - O título da tabela.
 * @param {Object} dadosCruzados - Os dados da análise.
 * @param {string} cabecalhoChave - O nome do cabeçalho da coluna principal.
 * @param {string} cabecalhoValor - O nome do cabeçalho da coluna secundária.
 */
function escreverTabelaCruzadaGeneric(aba, linhaInicio, titulo, dadosCruzados, cabecalhoChave, cabecalhoValor) {
  aba.getRange(linhaInicio, 1).setValue(titulo).setFontWeight('bold').setFontSize(14);
  let linhaAtual = linhaInicio + 2;

  aba.getRange(linhaAtual, 1, 1, 3).setValues([[cabecalhoChave, cabecalhoValor, 'Contagem']]).setFontWeight('bold');
  const dadosParaEscrever = [];

  for (const chave in dadosCruzados) {
    const valoresOrdenados = Object.entries(dadosCruzados[chave]).sort((a, b) => b[1] - a[1]);
    
    valoresOrdenados.forEach(parValor => {
      const valor = parValor[0];
      const contagem = parValor[1];
      dadosParaEscrever.push([chave, valor, contagem]);
    });
  }
  
  if (dadosParaEscrever.length > 0) {
    aba.getRange(linhaAtual + 1, 1, dadosParaEscrever.length, 3).setValues(dadosParaEscrever);
  }

  aba.autoResizeColumns(1, 3);
}
