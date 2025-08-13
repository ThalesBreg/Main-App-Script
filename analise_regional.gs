/**
 *
 * Este script continua a análise diagnóstica, cruzando os dados
 * de 'Gatilho' com 'Regional' para identificar concentrações
 * geográficas de problemas.
 */
move file to src/analise_regional.gs
// =================== CONFIGURAÇÕES ===================
const ID_PLANILHA_ANALISE_REG = '1nINa_O3EFr9GugHc3SIch2JfeMkWGgIJzVdKTeWSSTQ'; // ID da planilha de Destino/Análise
const NOME_ABA_DADOS_REG = 'bd_gcr_script';
const NOME_ABA_RESULTADO_REG = 'Analise_Diagnostica';

/**
 * Função principal para orquestrar a análise Gatilho vs. Regional.
 */
function executarAnaliseGatilhoVsRegional() {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_ANALISE_REG);
    const abaDados = planilha.getSheetByName(NOME_ABA_DADOS_REG);

    if (!abaDados) {
      throw new Error(`A aba de dados "${NOME_ABA_DADOS_REG}" não foi encontrada.`);
    }

    const dadosCompletos = abaDados.getDataRange().getValues();
    if (dadosCompletos.length < 2) {
      Logger.log('Não há dados suficientes para a análise regional.');
      return;
    }

    const cabecalhos = dadosCompletos.shift();

    // --- Mapeia os índices das colunas ---
    const indices = {};
    cabecalhos.forEach((col, i) => { indices[col] = i; });
    
    const indiceGatilho = indices['Gatilho'];
    const indiceRegional = indices['Regional'];

    if (indiceGatilho === undefined || indiceRegional === undefined) {
        throw new Error("As colunas 'Gatilho' e/ou 'Regional' não foram encontradas.");
    }

    // --- Realiza a contagem cruzada ---
    const analiseCruzada = contarAnaliseCruzadaGeneric(dadosCompletos, indiceGatilho, indiceRegional);

    // --- Prepara a aba de resultados ---
    let abaResultado = planilha.getSheetByName(NOME_ABA_RESULTADO_REG);
    if (!abaResultado) {
      // Se a aba não existir, avisa o usuário para rodar a análise anterior primeiro.
      throw new Error(`A aba de resultados "${NOME_ABA_RESULTADO_REG}" não foi encontrada. Por favor, execute a "Análise Diagnóstica - Gatilho vs Causal" primeiro.`);
    }

    // --- Encontra a última linha para adicionar a nova tabela ---
    const proximaLinha = abaResultado.getLastRow() + 3; // Adiciona um espaço de 2 linhas

    // --- Escreve os resultados na aba ---
    escreverTabelaCruzadaGeneric(abaResultado, proximaLinha, 'Análise Cruzada: Gatilho vs. Regional', analiseCruzada, 'Gatilho', 'Regional');

    SpreadsheetApp.getUi().alert('Sucesso!', `A análise Gatilho vs. Regional foi concluída e adicionada à aba "${NOME_ABA_RESULTADO_REG}".`, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Análise Gatilho vs. Regional concluída com sucesso.');

  } catch (e) {
    Logger.log(`Ocorreu um erro na análise regional: ${e.message}\nStack: ${e.stack}`);
    SpreadsheetApp.getUi().alert('Erro', `Ocorreu um erro: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Conta a frequência de uma coluna de valor dentro de uma coluna de chave.
 * @param {Array<Array<string>>} dados - Os dados da planilha.
 * @param {number} indiceChave - O índice da coluna principal (ex: Gatilho).
 * @param {number} indiceValor - O índice da coluna secundária (ex: Regional).
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
  // Título geral
  aba.getRange(linhaInicio, 1).setValue(titulo).setFontWeight('bold').setFontSize(14);
  let linhaAtual = linhaInicio + 2;

  // Cabeçalhos da tabela
  aba.getRange(linhaAtual, 1, 1, 3).setValues([[cabecalhoChave, cabecalhoValor, 'Contagem']]).setFontWeight('bold');
  const dadosParaEscrever = [];

  // Transforma o objeto aninhado em um array plano para escrita
  for (const chave in dadosCruzados) {
    const valoresOrdenados = Object.entries(dadosCruzados[chave]).sort((a, b) => b[1] - a[1]);
    
    valoresOrdenados.forEach(parValor => {
      const valor = parValor[0];
      const contagem = parValor[1];
      dadosParaEscrever.push([chave, valor, contagem]);
    });
  }
  
  // Escreve os dados na planilha
  if (dadosParaEscrever.length > 0) {
    aba.getRange(linhaAtual + 1, 1, dadosParaEscrever.length, 3).setValues(dadosParaEscrever);
  }

  // Formatação final
  aba.autoResizeColumns(1, 3);
}
