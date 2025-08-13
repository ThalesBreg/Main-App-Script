// Controle de envio: true = envia para consultores, false = só para o webhook padrão
var ENVIAR_PARA_CONSULTORES = true;
// Função para remover acentos e normalizar nomes
function removerAcentos(str) {
  if (!str) return '';
  return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '');
}

function normalizarNome(nome) {
  if (!nome) return '';
  return removerAcentos(nome.toLowerCase().trim().replace(/\s+/g, ''));
}
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

/**
 * Envia um resumo diário dos eventos dos Places de cada consultor via webhook.
 * Para cada consultor, lista os Places da sua carteira e os eventos do dia.
 */
function enviarResumoDiarioConsultor() {
    // Log detalhado de atribuição de consultor para cada caso
    const idxIdChamado = indices['ID do chamado'] !== undefined ? indices['ID do chamado'] :
                         (indices['ID_Chamado'] !== undefined ? indices['ID_Chamado'] :
                         (indices['ID CHAMADO'] !== undefined ? indices['ID CHAMADO'] : null));
    dadosCompletos.forEach(linha => {
      const idChamado = idxIdChamado !== null ? linha[idxIdChamado] : '';
      const consultor = (linha[indices['CONSULTOR']] || 'Não informado').trim();
      const consultorNormalizado = normalizarNome(consultor);
      Logger.log(`Auditoria: Chamado ${idChamado} | Consultor atribuído: ${consultor} | Normalizado: ${consultorNormalizado}`);
    });
  try {
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

    const hoje = new Date();
    hoje.setHours(0,0,0,0);

    // Agrupar eventos do dia por consultor
    const eventosPorConsultor = {};
    dadosCompletos.forEach(linha => {
      const dataChamado = new Date(linha[indices['Carimbo de data/hora']]);
      dataChamado.setHours(0,0,0,0);
      if (dataChamado.getTime() === hoje.getTime()) {
        const consultor = (linha[indices['CONSULTOR']] || 'Não informado').trim();
        if (!eventosPorConsultor[consultor]) eventosPorConsultor[consultor] = [];
        eventosPorConsultor[consultor].push(linha);
      }
    });

    if (ENVIAR_PARA_CONSULTORES) {
      // Enviar resumo simples para cada consultor
      Object.keys(eventosPorConsultor).forEach(consultor => {
        const eventos = eventosPorConsultor[consultor];
        const nomeNormalizado = normalizarNome(consultor);
        Logger.log(`Consultor original: ${consultor} | Normalizado: ${nomeNormalizado}`);
        const webhookUrl = consultorWebhooks[nomeNormalizado];
        if (webhookUrl && !webhookUrl.includes('COLE_A_URL')) {
          // Agrupar por gatilho
          const gatilhos = {};
          // Buscar índice do ID do chamado considerando variações
          const idxIdChamado = indices['ID do chamado'] !== undefined ? indices['ID do chamado'] :
                               (indices['ID_Chamado'] !== undefined ? indices['ID_Chamado'] :
                               (indices['ID CHAMADO'] !== undefined ? indices['ID CHAMADO'] : null));
          eventos.forEach(linha => {
            const gatilho = linha[indices['Gatilho']];
            const placeId = linha[indices['Place ID Tratado']];
            const chamadoId = idxIdChamado !== null ? linha[idxIdChamado] : '';
            if (!gatilhos[gatilho]) gatilhos[gatilho] = [];
            gatilhos[gatilho].push({ placeId, chamadoId });
          });
          let mensagem = `Olá, ${consultor},\n\nTivemos ${eventos.length} chamados na sua carteira hoje, sobre esses assuntos:\n`;
          Object.keys(gatilhos).forEach(gatilho => {
            mensagem += `\n  • ${gatilho}\n`;
            gatilhos[gatilho].forEach(item => {
              mensagem += `    * Place: ${item.placeId} | Chamado: ${item.chamadoId}\n`;
            });
          });
          mensagem += `\n---\nEste é um resumo automático do sistema de alertas.`;
          try {
            const response = UrlFetchApp.fetch(webhookUrl, {
              'method': 'post',
              'contentType': 'application/json',
              'payload': JSON.stringify({ 'text': mensagem })
            });
            Logger.log(`Resumo enviado para consultor: ${consultor} | Webhook: ${webhookUrl} | Status: ${response.getResponseCode()}`);
          } catch(e) {
            Logger.log(`Falha no envio do resumo para ${consultor}: ${e.message}`);
          }
        } else {
          Logger.log(`Resumo NÃO enviado para consultor: ${consultor} (sem webhook configurado)`);
        }
      });
    } else {
      // Envia um resumo geral para o webhook padrão
      let mensagemGeral = 'Resumo diário dos consultores:\n';
      Object.keys(eventosPorConsultor).forEach(consultor => {
        const eventos = eventosPorConsultor[consultor];
        // Agrupar por gatilho
        const gatilhos = {};
        const idxIdChamado = indices['ID do chamado'] !== undefined ? indices['ID do chamado'] :
                             (indices['ID_Chamado'] !== undefined ? indices['ID_Chamado'] :
                             (indices['ID CHAMADO'] !== undefined ? indices['ID CHAMADO'] : null));
        eventos.forEach(linha => {
          const gatilho = linha[indices['Gatilho']];
          const placeId = linha[indices['Place ID Tratado']];
          const chamadoId = idxIdChamado !== null ? linha[idxIdChamado] : '';
          if (!gatilhos[gatilho]) gatilhos[gatilho] = [];
          gatilhos[gatilho].push({ placeId, chamadoId });
        });
        mensagemGeral += `\nConsultor: ${consultor} (${eventos.length} chamados)\n`;
        Object.keys(gatilhos).forEach(gatilho => {
          mensagemGeral += `  • ${gatilho}\n`;
          gatilhos[gatilho].forEach(item => {
            mensagemGeral += `    * Place: ${item.placeId} | Chamado: ${item.chamadoId}\n`;
          });
        });
      });
      mensagemGeral += '\n---\nEste é um resumo automático do sistema de alertas.';
      try {
        const response = UrlFetchApp.fetch(WEBHOOK_PADRAO, {
          'method': 'post',
          'contentType': 'application/json',
          'payload': JSON.stringify({ 'text': mensagemGeral })
        });
        Logger.log(`Resumo diário enviado para o webhook padrão. Status: ${response.getResponseCode()}`);
      } catch(e) {
        Logger.log(`Falha no envio do resumo diário para o webhook padrão: ${e.message}`);
      }
    }
  } catch (e) {
    Logger.log(`Erro ao enviar resumo diário: ${e.message}`);
  }
}

/**
 * Alerta imediato de pausa de Place: envia alerta via webhook para o consultor responsável
 * sempre que um chamado de pausa for aberto na data atual.
 * Pode ser executada por acionador (trigger) diário ou a cada hora.
 */
function alertaPausaPlaceHoje() {
  try {
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

    const hoje = new Date();
    hoje.setHours(0,0,0,0);

    const chamadosPausaHoje = dadosCompletos.filter(linha => {
      const dataChamado = new Date(linha[indices['Carimbo de data/hora']]);
      dataChamado.setHours(0,0,0,0);
      return linha[indices['Gatilho']] === 'PAUSA' && dataChamado.getTime() === hoje.getTime();
    });

    if (chamadosPausaHoje.length > 0) {
      Logger.log(`Encontrados ${chamadosPausaHoje.length} chamados de pausa para hoje.`);
      chamadosPausaHoje.forEach(chamado => {
        const placeId = chamado[indices['Place ID Tratado']];
        const consultor = chamado[indices['CONSULTOR']] || 'Não informado';
        const webhookUrl = consultorWebhooks[consultor.trim()] || DEFAULT_WEBHOOK_URL;
        if (webhookUrl && !webhookUrl.includes('COLE_A_URL')) {
          const mensagem = `*⏸️ Alerta de Pausa de Place*\n\n` +
                          `O Place ID: ${placeId} foi pausado hoje.\n` +
                          `Consultor responsável: ${consultor}\n` +
                          `_Acompanhe e atue conforme necessário._`;
          try {
            UrlFetchApp.fetch(webhookUrl, {
              'method': 'post',
              'contentType': 'application/json',
              'payload': JSON.stringify({ 'text': mensagem })
            });
          } catch(e) {
            Logger.log(`Falha no envio via Workchat para ${placeId}: ${e.message}`);
          }
        } else {
          Logger.log(`Webhook não configurado para "${consultor}".`);
        }
      });
    } else {
      Logger.log('Nenhum chamado de pausa encontrado para hoje.');
    }
  } catch (e) {
    Logger.log(`Erro ao enviar alerta de pausa: ${e.message}`);
  }
}

function carregarConfiguracaoWebhooks(abaConfig) {
  const dadosConfig = abaConfig.getDataRange().getValues();
  dadosConfig.shift();
  const webhooks = {};
  dadosConfig.forEach(linha => {
    const nome = linha[0];
    const url = linha[1];
    if (nome && url) {
      webhooks[nome.trim()] = url.trim();
      nomesConfigLog.push(nome);
    }
  });
  Logger.log(`${Object.keys(webhooks).length} configurações de webhook carregadas.`);
  Logger.log(`Nomes normalizados na configuração: ${JSON.stringify(nomesConfigLog)}`);
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
      let webhookUrl = consultorWebhooks[consultor];
      if (!ENVIAR_PARA_CONSULTORES) {
        webhookUrl = DEFAULT_WEBHOOK_URL;
      }
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

const nomesConfigLog = [];
