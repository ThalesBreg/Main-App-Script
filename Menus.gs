/**
 *
 * Este script cria um menu personalizado na interface da Planilha Google
 * para facilitar a execução das funções principais do projeto.
 * Ele é executado automaticamente sempre que a planilha é aberta.
 */
function onOpen() {
move file to src/Menus.gs
      .createMenu('🤖 Automações Meli')
      .addItem('1. Processar Novos Dados (ETL)', 'processarDadosDaOrigemParaDestino')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('2. Análises')
          .addItem('Executar Análise Estatística', 'executarAnaliseEstatistica')
          .addItem('Executar Análise Diagnóstica', 'executarAnaliseDiagnostica')
          .addItem('Executar Análise Temporal', 'executarAnaliseTemporal'))
      .addSeparator()
      .addItem('3. Verificar Risco de Churn (Manual)', 'verificarRiscoDeChurn_v4')
      .addToUi();
}
