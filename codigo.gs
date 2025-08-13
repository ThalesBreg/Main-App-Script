/**
 * Este script foi projetado para automatizar a coleta, limpeza e enriquecimento
 * de dados de planilhas de origem para uma planilha de banco de dados central.
 * As configurações são lidas do arquivo 'Configuracoes.gs'.
 */

/**
 * Ponto de entrada para processar os dados de GCR.
 * Orquestra a cópia, processamento e enriquecimento dos dados do formulário.
move file to src/codigo.gs
function processarDadosDaOrigemParaDestino() {
  try {
    // As constantes como ID_PLANILHA_PRINCIPAL são lidas do arquivo Configuracoes.gs
    const planilhaDestino = SpreadsheetApp.openById(ID_PLANILHA_PRINCIPAL);

    // --- PASSO 1: Carregar a lista de Places com pendências de PEP ---
    const abaPep = planilhaDestino.getSheetByName(NOME_ABA_PEP);
    if (!abaPep) {
      throw new Error(`Aba de referência de PEP "${NOME_ABA_PEP}" não foi encontrada. Por favor, crie-a.`);
    }
    const ultimaLinhaPep = abaPep.getLastRow();
    let placesComPep = new Set();
    if (ultimaLinhaPep > 1) {
      const dadosPep = abaPep.getRange(2, 1, ultimaLinhaPep - 1, 1).getValues();
      placesComPep = new Set(dadosPep.map(linha => String(linha[0])));
    }
    Logger.log(`${placesComPep.size} Places com pendência de PEP carregados.`);

    // --- PASSO 2: Ler os dados brutos da origem GCR ---
    const planilhaOrigem = SpreadsheetApp.openById(ID_PLANILHA_ORIGEM_GCR);
    const abaOrigem = planilhaOrigem.getSheetByName(NOME_ABA_ORIGEM_GCR);
    if (!abaOrigem) {
      throw new Error(`Aba de origem "${NOME_ABA_ORIGEM_GCR}" não encontrada.`);
    }
    const ultimaLinhaOrigem = abaOrigem.getLastRow();
    if (ultimaLinhaOrigem < 2) {
      Logger.log('Não há dados para processar além do cabeçalho na origem GCR.');
      return;
    }
    const NUMERO_DE_COLUNAS_ORIGEM = 16;
    const dadosBrutos = abaOrigem.getRange(2, 1, ultimaLinhaOrigem - 1, NUMERO_DE_COLUNAS_ORIGEM).getValues();

    // --- PASSO 3: Processamento inicial e enriquecimento básico dos dados GCR ---
    let dadosProcessados = dadosBrutos
      .filter(linha => linha[0] && linha[0] instanceof Date)
      .map(linha => {
        const data = linha[0];
        const nomeMes = data.toLocaleDateString('pt-BR', { month: 'long', timeZone: 'UTC' });
        const numeroSemana = obterNumeroDaSemana(data);
        const placeIdOriginal = String(linha[7]);
        const placeIdTratado = placeIdOriginal.replace('849817033_', '');
        const isFinanceiro = String(linha[3]).startsWith('02.');
        const classificacao = isFinanceiro ? 'Financeiro' : 'Outros';
        const temPendenciaPepNaBase = placesComPep.has(placeIdTratado);
        const statusPep = temPendenciaPepNaBase && isFinanceiro ? 'Com Pendência PEP' : 'OK';

        const linhaModificada = [
            ...linha.slice(0, 5), classificacao, statusPep,
            ...linha.slice(5, 8), placeIdTratado, ...linha.slice(8)
        ];
        linhaModificada.push(nomeMes, numeroSemana);
        return linhaModificada;
      });

    // --- PASSO 4: Carregar dados da Carteira para cruzamento ---
    const mapaCarteira = criarMapaDeReferenciaCarteira(planilhaDestino);
    const COLUNAS_DA_CARTEIRA = ["STATUS PORTAL", "SUB-REGIONAL", "CONSULTOR", "TEAM LEADER", "PEP", "LATITUDE", "LONGITUDE", "SVC"];

    // --- PASSO 5: Enriquecer dados do GCR com dados da Carteira e criar coluna LAT_LONG ---
    const CABECALHOS_INICIAIS = [
      "Carimbo de data/hora", "Email", "ID_Chamado", "Gatilho", "Causal", "Classificação", "Status PEP",
      "Serviço", "Status", "Place ID", "Place ID Tratado", "Descrição", "Evidências",
      "Regional", "Cluster", "PLC_PLACE_SVC", "PLACE_NAME", "Andamento", "Resolução", "Mês", "Semana do Ano"
    ];
    const CABECALHOS_FINAIS = [...CABECALHOS_INICIAIS, ...COLUNAS_DA_CARTEIRA, "LAT_LONG"];
    const indicePlaceIdTratado = 10; 

    const dadosFinais = dadosProcessados.map(linhaGCR => {
        const placeId = linhaGCR[indicePlaceIdTratado];
        const infoCarteira = mapaCarteira.get(placeId) || {};
        const dadosExtras = COLUNAS_DA_CARTEIRA.map(coluna => infoCarteira[coluna] || ""); 
        const latitude = infoCarteira["LATITUDE"];
        const longitude = infoCarteira["LONGITUDE"];
        let latLongCombinado = (latitude && longitude) ? `${latitude},${longitude}` : "";
        return [...linhaGCR, ...dadosExtras, latLongCombinado];
    });

    // --- PASSO 6: Ordenar e escrever os dados na planilha de destino ---
    dadosFinais.sort((a, b) => b[0].getTime() - a[0].getTime());

    if (dadosFinais.length === 0) {
      Logger.log('Nenhum dado válido encontrado para copiar.');
      return;
    }

    let abaDestino = planilhaDestino.getSheetByName(NOME_ABA_DADOS_GCR);
    if (!abaDestino) {
      abaDestino = planilhaDestino.insertSheet(NOME_ABA_DADOS_GCR);
    }

    abaDestino.clear();
    abaDestino.getRange(1, 1, 1, CABECALHOS_FINAIS.length).setValues([CABECALHOS_FINAIS]);
    abaDestino.getRange(2, 1, dadosFinais.length, dadosFinais[0].length).setValues(dadosFinais);
    abaDestino.setFrozenRows(1);
    abaDestino.getRange(1, 1, 1, CABECALHOS_FINAIS.length).setFontWeight('bold');

    Logger.log(`Sucesso! ${dadosFinais.length} linhas de GCR foram processadas e salvas na aba "${NOME_ABA_DADOS_GCR}".`);

  } catch (e) {
    Logger.log(`Ocorreu um erro crítico no processamento de GCR: ${e.message}\nStack: ${e.stack}`);
  }
}

/**
 * Lê a aba da carteira e cria um mapa para consulta rápida.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} planilhaDestino O objeto da planilha de destino.
 * @returns {Map<string, Object>} Um mapa com o Place ID como chave e as informações da linha como valor.
 */
function criarMapaDeReferenciaCarteira(planilhaDestino) {
    // VERIFICAÇÃO: Garante que a função não foi chamada diretamente sem o parâmetro necessário.
    if (!planilhaDestino) {
        throw new Error("A função 'criarMapaDeReferenciaCarteira' foi chamada sem o objeto da planilha. " +
                        "Certifique-se de executar a função principal 'processarDadosDaOrigemParaDestino' em vez desta.");
    }

    const abaCarteira = planilhaDestino.getSheetByName(NOME_ABA_DESTINO_CARTEIRA);
    if (!abaCarteira) {
        Logger.log(`Aba de referência da carteira "${NOME_ABA_DESTINO_CARTEIRA}" não encontrada. O enriquecimento será pulado.`);
        return new Map();
    }
    const dadosCarteira = abaCarteira.getDataRange().getValues();
    const cabecalhos = dadosCarteira.shift(); 
    
    const indices = {};
    cabecalhos.forEach((col, i) => indices[col] = i);

    const indicePlaceID = indices["PLACE ID"];
    if (indicePlaceID === undefined) {
        throw new Error("A coluna 'PLACE ID' não foi encontrada na aba da carteira.");
    }

    const mapa = new Map();
    dadosCarteira.forEach(linha => {
        const placeId = String(linha[indicePlaceID]);
        if (placeId) {
            const info = {};
            cabecalhos.forEach((coluna, i) => {
                info[coluna] = linha[i];
            });
            mapa.set(placeId, info);
        }
    });
    Logger.log(`${mapa.size} registros da carteira carregados para cruzamento.`);
    return mapa;
}

/**
 * Calcula o número da semana para uma data específica.
 * @param {Date} data A data para a qual o número da semana será calculado.
 * @returns {number} O número da semana no ano.
 */
function obterNumeroDaSemana(data) {
  const d = new Date(Date.UTC(data.getFullYear(), data.getMonth(), data.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const anoInicio = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const numeroSemana = Math.ceil((((d - anoInicio) / 86400000) + 1) / 7);
  return numeroSemana;
}

/**
 * Ponto de entrada para processar os dados da Carteira.
 */
function processarDadosCarteira() {
  try {
    const planilhaOrigem = SpreadsheetApp.openById(ID_PLANILHA_ORIGEM_CARTEIRA);
    const abaOrigem = planilhaOrigem.getSheetByName(NOME_ABA_ORIGEM_CARTEIRA);

    if (!abaOrigem) {
      throw new Error(`Aba de origem da carteira "${NOME_ABA_ORIGEM_CARTEIRA}" não encontrada.`);
    }

    const dadosCompletos = abaOrigem.getDataRange().getValues();

    if (dadosCompletos.length < 2) {
      Logger.log('Não há dados para processar na base de carteira.');
      return;
    }

    const planilhaDestino = SpreadsheetApp.openById(ID_PLANILHA_PRINCIPAL);
    let abaDestino = planilhaDestino.getSheetByName(NOME_ABA_DESTINO_CARTEIRA);
    if (!abaDestino) {
      abaDestino = planilhaDestino.insertSheet(NOME_ABA_DESTINO_CARTEIRA);
    }

    abaDestino.clear();
    abaDestino.getRange(1, 1, dadosCompletos.length, dadosCompletos[0].length).setValues(dadosCompletos);
    abaDestino.setFrozenRows(1);
    abaDestino.getRange(1, 1, 1, dadosCompletos[0].length).setFontWeight('bold');

    Logger.log(`Sucesso! ${dadosCompletos.length - 1} linhas da carteira foram salvas na aba "${NOME_ABA_DESTINO_CARTEIRA}".`);

  } catch (e) {
    Logger.log(`Ocorreu um erro crítico no processamento da Carteira: ${e.message}\nStack: ${e.stack}`);
  }
}
