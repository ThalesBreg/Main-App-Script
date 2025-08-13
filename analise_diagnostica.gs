/**
 *
 * Este script realiza uma análise diagnóstica, cruzando os dados
 * de 'Gatilho' com 'Causal' para identificar as causas mais comuns.
 * As configurações são lidas do arquivo 'Configuracoes.gs'.
 */
move file to src/analise_diagnostica.gs
/**
 * Função principal para orquestrar a análise diagnóstica.
 */
function executarAnaliseDiagnostica() {
  try {
    // As constantes são lidas do arquivo Configuracoes.gs
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_PRINCIPAL);
    const abaDados = planilha.getSheetByName(NOME_ABA_DADOS_GCR);

    if (!abaDados) {
      throw new Error(`A aba de dados "${NOME_ABA_DADOS_GCR}" não foi encontrada.`);
    }

    const dadosCompletos = abaDados.getDataRange().getValues();
    if (dadosCompletos.length < 2) {
      Logger.log('Não há dados suficientes para a análise diagnóstica.');
      return;
    }

    const cabecalhos = dadosCompletos.shift();

    const indices = {};
    cabecalhos.forEach((col, i) => { indices[col] = i; });
    
    const indiceGatilho = indices['Gatilho'];
    const indiceCausal = indices['Causal'];

    if (indiceGatilho === undefined || indiceCausal === undefined) {
        throw new Error("As colunas 'Gatilho' e/ou 'Causal' não foram encontradas.");
    }

    const analiseCruzada = contarGatilhoVsCausal(dadosCompletos, indiceGatilho, indiceCausal);

    let abaResultado = planilha.getSheetByName(NOME_ABA_ANALISE_DIAGNOSTICA);
    if (abaResultado) {
      abaResultado.clear();
    } else {
      abaResultado = planilha.insertSheet(NOME_ABA_ANALISE_DIAGNOSTICA);
    }

    escreverTabelaCruzada(abaResultado, 'Análise Cruzada: Gatilho vs. Causal', analiseCruzada);

    SpreadsheetApp.getUi().alert('Sucesso!', `A análise diagnóstica foi concluída e os resultados estão na aba "${NOME_ABA_ANALISE_DIAGNOSTICA}".`, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Análise diagnóstica concluída com sucesso.');

  } catch (e) {
    Logger.log(`Ocorreu um erro na análise diagnóstica: ${e.message}\nStack: ${e.stack}`);
    SpreadsheetApp.getUi().alert('Erro', `Ocorreu um erro: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Conta a frequência de cada 'Causal' dentro de cada 'Gatilho'.
 */
function contarGatilhoVsCausal(dados, indiceGatilho, indiceCausal) {
  const contagens = {};
  dados.forEach(linha => {
    const gatilho = linha[indiceGatilho] || 'Vazio';
    const causal = linha[indiceCausal] || 'Vazio';

    if (!contagens[gatilho]) {
      contagens[gatilho] = {};
    }
    contagens[gatilho][causal] = (contagens[gatilho][causal] || 0) + 1;
  });
  return contagens;
}

/**
 * Escreve a tabela de análise cruzada na aba de resultados.
 */
function escreverTabelaCruzada(aba, titulo, dadosCruzados) {
  let linhaAtual = 1;

  aba.getRange(linhaAtual, 1).setValue(titulo).setFontWeight('bold').setFontSize(14);
  linhaAtual += 2;

  aba.getRange(linhaAtual, 1, 1, 3).setValues([['Gatilho', 'Causal', 'Contagem']]).setFontWeight('bold');
  const dadosParaEscrever = [];

  for (const gatilho in dadosCruzados) {
    const causaisOrdenados = Object.entries(dadosCruzados[gatilho]).sort((a, b) => b[1] - a[1]);
    
    causaisOrdenados.forEach(parCausal => {
      const causal = parCausal[0];
      const contagem = parCausal[1];
      dadosParaEscrever.push([gatilho, causal, contagem]);
    });
  }
  
  if (dadosParaEscrever.length > 0) {
    aba.getRange(linhaAtual + 1, 1, dadosParaEscrever.length, 3).setValues(dadosParaEscrever);
  }

  aba.autoResizeColumns(1, 3);
  aba.setFrozenRows(linhaAtual);
}
