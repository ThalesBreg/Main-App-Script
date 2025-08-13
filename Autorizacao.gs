/**
 * FUNÇÃO DE AUTORIZAÇÃO
 *
 * O único propósito desta função é forçar a janela de autorização do Google a aparecer.
 * Execute esta função UMA VEZ para dar ao projeto as permissões necessárias para
 * ler e escrever em planilhas.
 */
function solicitarAutorizacao() {
  try {
move file to src/Autorizacao.gs
    // A constante ID_PLANILHA_PRINCIPAL é lida do arquivo Configuracoes.gs
    SpreadsheetApp.openById(ID_PLANILHA_PRINCIPAL);
    
    // Se o código chegar aqui, significa que as permissões já foram dadas.
    SpreadsheetApp.getUi().alert('Sucesso!', 'As permissões foram concedidas corretamente. Agora você pode usar todas as outras funções.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // Se a autorização ainda não foi dada, um erro de permissão será capturado aqui,
    // mas a janela de autorização já terá sido exibida para o usuário.
    SpreadsheetApp.getUi().alert('Ação Necessária', 'A janela de autorização foi exibida. Por favor, complete o processo de permissão e, em seguida, execute a função desejada novamente.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
