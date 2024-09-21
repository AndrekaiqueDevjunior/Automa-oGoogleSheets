function configurarCabecalhoHistorico() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var abaHistorico = spreadsheet.getSheetByName('Histórico');
  
  if (!abaHistorico) {
    SpreadsheetApp.getUi().alert('A aba "Histórico" não foi encontrada.');
    return;
  }
  
  // Definir os títulos do cabeçalho
  var cabecalho = [
    'Data', 'Número da Proposta', 'Cliente', 'CNPJ', 'Comprador', 'Email do Comprador',
    'Telefone', 'WhatsApp', 'Endereço', 'Número', 'Complemento', 'Bairro', 'Cidade',
    'Estado', 'Telefone Geral', 'Email Geral', 'Data da Proposta', 'Validade da Proposta',
    'Condição de Pagamento', 'Frete', 'Transportadora', 'Valor Total da Proposta', 'Impostos',
    'Frete Total', 'Vendedor', 'Email do Vendedor', 'Telefone do Vendedor', 'WhatsApp do Vendedor',
    'Item', 'Fabricante', 'Código', 'Descrição', 'NCM', 'Quantidade', 'Prazo de Entrega',
    'Quantidade em Estoque', 'Valor Unitário', 'Desconto', 'Valor Unitário com Desconto',
    'Valor Total do Item', 'Link da Proposta'
  ];
  
  // Definir o cabeçalho na primeira linha
  abaHistorico.getRange('A1:AO1').setValues([cabecalho]);
  
  SpreadsheetApp.getUi().alert('Cabeçalho da aba "Histórico" configurado com sucesso.');
}
