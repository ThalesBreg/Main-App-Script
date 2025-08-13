/**
 * SISTEMA DE ALERTA DE RISCO DE CHURN
 *
 * Este script utiliza as variáveis definidas no arquivo 'Configuracoes.gs'.
 * Ele monitora chamados de risco, enriquece o alerta com o histórico do parceiro
 * e envia notificações direcionadas.
 */

/**
move file to src/Alerta.gs
 * Deve ser executada por um acionador (trigger) baseado em tempo.
 */
function verificarRiscoDeChurn_v4() {
  try {
    // As variáveis de configuração são lidas diretamente do arquivo Configuracoes.gs
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_PRINCIPAL);
    const abaDados = planilha.getSheetByName(NOME_ABA_DADOS_GCR);
    const abaConfig = planilha.getSheetByName(NOME_ABA_CONFIG_CONSULTORES);

    if (!abaDados) throw new Error(`Aba de dados "${NOME_ABA_DADOS_GCR}" não foi encontrada.`);
    if (!abaConfig) throw new Error(`Aba de configuração "${NOME_ABA_CONFIG_CONSULTORES}" não foi encontrada.`);

    const consultorWebhooks = carregarConfiguracaoWebhooks(abaConfig);
    const dadosCompletos = abaDados.getDataRange().getValues();
    const cabecalhos = dadosCompletos.shift();
    const indices = mapearIndices(cabecalhos);
    validarColunas(indices);

    const scriptProperties = PropertiesService.getScriptProperties();
    const ultimaVerificacao = scriptProperties.getProperty('ultimaVerificacao_v4') || 0;
    const novaUltimaVerificacao = new Date().getTime();

    const chamadosDeRisco = dadosCompletos.filter(linha => {
      const timestampLinha = new Date(linha[indices['Carimbo de data/hora']]).getTime();
      return linha[indices['Gatilho']] === GATILHO_DE_RISCO &&
             linha[indices['Causal']] === CAUSAL_DE_RISCO &&
             timestampLinha > ultimaVerificacao;
    });

    if (chamadosDeRisco.length > 0) {
      Logger.log(`Encontrados ${chamadosDeRisco.length} novos chamados de risco.`);
      chamadosDeRisco.forEach(chamado => {
        const placeId = chamado[indices['Place ID Tratado']];
        const resumoHistorico = gerarResumoHistorico(placeId, dadosCompletos, indices);
        enviarAlertaMulticanal(chamado, indices, resumoHistorico, consultorWebhooks);
      });
    } else {
      Logger.log('Nenhum novo chamado de risco de churn encontrado.');
    }

    scriptProperties.setProperty('ultimaVerificacao_v4', novaUltimaVerificacao);

  } catch (e) {
    Logger.log(`Erro ao verificar risco de churn: ${e.message}`);
  }
}

// --- FUNÇÕES DE LÓGICA (Não precisam de alteração) ---

function carregarConfiguracaoWebhooks(abaConfig) {
  const dadosConfig = abaConfig.getDataRange().getValues();
  dadosConfig.shift();
  const webhooks = {};
  dadosConfig.forEach(linha => {
    const nome = linha[0];
    const url = linha[1];
    if (nome && url) {
      webhooks[nome.trim()] = url.trim();
    }
  });
  Logger.log(`${Object.keys(webhooks).length} configurações de webhook carregadas.`);
  return webhooks;
}

function gerarResumoHistorico(placeId, dados, indices) {
  const chamadosDoPlace = dados.filter(linha => linha[indices['Place ID Tratado']] === placeId);
  const contagemPorGatilho = {};
  chamadosDoPlace.forEach(chamado => {
    const gatilho = chamado[indices['Gatilho']] || 'Não especificado';
    contagemPorGatilho[gatilho] = (contagemPorGatilho[gatilho] || 0) + 1;
  });
  let resumoFormatado = '';
  for (const gatilho in contagemPorGatilho) {
    resumoFormatado += `\n  • ${contagemPorGatilho[gatilho]}x - ${gatilho}`;
  }
  return { total: chamadosDoPlace.length, resumo: resumoFormatado };
}

function enviarAlertaMulticanal(chamado, indices, resumoHistorico, consultorWebhooks) {
  const placeId = chamado[indices['Place ID Tratado']];
  const dataChamado = new Date(chamado[indices['Carimbo de data/hora']]).toLocaleString('pt-BR');
  const regional = chamado[indices['Regional']] || 'Não informado';
  const consultor = chamado[indices['CONSULTOR']] || 'Não informado';
  const assunto = `ALERTA DE RISCO DE CHURN: Place ${placeId} solicitou pausa`;

  try {
    const corpoEmail = `
      <html><body>
        <h2>🚨 Alerta Proativo de Risco de Churn</h2>
        <p>Um parceiro sinalizou intenção de pausar as atividades. Recomenda-se contato imediato.</p><hr>
        <h3>Detalhes do Parceiro:</h3>
        <ul>
          <li><strong>Place ID:</strong> ${placeId}</li>
          <li><strong>Consultor:</strong> ${consultor}</li>
          <li><strong>Regional:</strong> ${regional}</li>
          <li><strong>Total de Chamados Históricos:</strong> ${resumoHistorico.total}</li>
        </ul>
        <h4>Resumo do Histórico:</h4>
        <pre>${resumoHistorico.resumo.replace(/\n/g, '<br>')}</pre><hr>
        <p><em>Este é um e-mail automático do Sistema de Monitoramento de Risco.</em></p>
      </body></html>`;
    MailApp.sendEmail({ to: EMAIL_DESTINATARIO_ALERTA, subject: assunto, htmlBody: corpoEmail, name: 'Monitor de Risco Places' });
  } catch(e) { Logger.log(`Falha no envio do e-mail para ${placeId}: ${e.message}`); }

  try {
    const webhookUrl = consultorWebhooks[consultor.trim()] || DEFAULT_WEBHOOK_URL;
    if (webhookUrl && !webhookUrl.includes('COLE_A_URL')) {
      const mensagemWorkchat = `*🚨 Alerta de Risco de Churn 🚨*\n\n` +
                               `Um de seus parceiros (*Place ID: ${placeId}*) pediu pausa!\n\n` +
                               `*Histórico (${resumoHistorico.total} chamados):*` +
                               `${resumoHistorico.resumo}\n\n` +
                               `_Ação imediata recomendada._`;
      UrlFetchApp.fetch(webhookUrl, {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify({ 'text': mensagemWorkchat })
      });
    } else {
      Logger.log(`Webhook não configurado para "${consultor}".`);
    }
  } catch(e) { Logger.log(`Falha no envio via Workchat para ${placeId}: ${e.message}`); }
}

function mapearIndices(cabecalhos) {
  const indices = {};
  cabecalhos.forEach((col, i) => { indices[col.trim()] = i; });
  return indices;
}

function validarColunas(indices) {
  const colunasNecessarias = ['Carimbo de data/hora', 'Gatilho', 'Causal', 'Place ID Tratado', 'Regional', 'CONSULTOR'];
  colunasNecessarias.forEach(coluna => {
    if (indices[coluna] === undefined) throw new Error(`Coluna "${coluna}" não encontrada.`);
  });
}
