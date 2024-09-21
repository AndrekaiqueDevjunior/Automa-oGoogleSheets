//function onOpen() {
  //var ui = SpreadsheetApp.getUi();
 /// ui.createMenu('Formulário')
  //  .addItem('Abrir Formulário', 'abrirFormulario')
 //   .addToUi();
//}

function abrirFormulario() {
  var html = HtmlService.createHtmlOutputFromFile('formulario') // formulario.html
      .setWidth(1000)
      .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Formulário de Dados');
}

// Função para gravar dados em uma aba específica
function gravarDados(dados) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaProposta = planilha.getSheetByName('FormularioProposta2024'); // Verifique o nome da aba

  // Log para verificar os dados recebidos
  Logger.log(dados);

  if (abaProposta) { // Verifica se a aba existe
    try {
      // Preenchendo células específicas
      abaProposta.getRange('A5').setValue(dados.cliente || '');
      abaProposta.getRange('C5').setValue(dados.nomeFantasia || '');
      abaProposta.getRange('E5').setValue(dados.cnpj || '');
      abaProposta.getRange('G5').setValue(dados.comprador || '');
      abaProposta.getRange('J5').setValue(dados.emailComprador || '');
      abaProposta.getRange('A7').setValue(dados.telefone || '');
      abaProposta.getRange('C7').setValue(dados.whatsapp || '');
      abaProposta.getRange('D7').setValue(dados.endereco || '');
      abaProposta.getRange('E7').setValue(dados.numero || '');
      abaProposta.getRange('F7').setValue(dados.complemento || '');
      abaProposta.getRange('G7').setValue(dados.bairro || '');
      abaProposta.getRange('I7').setValue(dados.cidade || '');
      abaProposta.getRange('J7').setValue(dados.estado || '');
      abaProposta.getRange('L7').setValue(dados.telefoneGeral || '');
      abaProposta.getRange('L7').setValue(dados.emailGeral || '');
      abaProposta.getRange('A9').setValue(dados.ad1 || '');
      abaProposta.getRange('B9').setValue(dados.ad2 || '');
      abaProposta.getRange('C9').setValue(dados.ad3 || '');
      abaProposta.getRange('D9').setValue(dados.ad4 || '');
      abaProposta.getRange('E9').setValue(dados.ad5 || '');
      abaProposta.getRange('F9').setValue(dados.ad6 || '');
      abaProposta.getRange('G9').setValue(dados.ad7 || '');
      abaProposta.getRange('I9').setValue(dados.ad8 || '');
      abaProposta.getRange('J9').setValue(dados.ad9 || '');
      abaProposta.getRange('K9').setValue(dados.ad10 || '');
      abaProposta.getRange('L9').setValue(dados.ad11 || '');
      abaProposta.getRange('A12').setValue(dados.dataProposta || '');
      abaProposta.getRange('F12').setValue(dados.validadeProposta || '');
      abaProposta.getRange('A15').setValue(dados.condicaoPagamento || '');
      abaProposta.getRange('F15').setValue(dados.frete || '');
      abaProposta.getRange('I15').setValue(dados.transportadora || '');
      abaProposta.getRange('G46').setValue(dados.valorTotalProposta || '');
      abaProposta.getRange('G47').setValue(dados.impostos || '');
      abaProposta.getRange('G48').setValue(dados.frete2 || '');
      abaProposta.getRange('G49').setValue(dados.notaPrazoEntrega || '');
      abaProposta.getRange('G51').setValue(dados.vendedor || '');
      abaProposta.getRange('G52').setValue(dados.emailVendedor || '');
      abaProposta.getRange('G53').setValue(dados.telefoneVendedor || '');
      abaProposta.getRange('G54').setValue(dados.whatsappVendedor || '');
      
      // Log para confirmar o preenchimento dos campos principais
      Logger.log("Dados gravados com sucesso nas células principais.");
      
      // Para os itens, você pode precisar de uma lógica para preencher várias linhas
      for (var i = 0; i < dados.itens.length; i++) {
        abaProposta.getRange(19 + i, 1).setValue(dados.itens[i].item || '');
        abaProposta.getRange(19 + i, 2).setValue(dados.itens[i].fabricante || '');
        abaProposta.getRange(19 + i, 3).setValue(dados.itens[i].codigo || '');
        abaProposta.getRange(19 + i, 4).setValue(dados.itens[i].descricao || '');
        abaProposta.getRange(19 + i, 5).setValue(dados.itens[i].ncm || '');
        abaProposta.getRange(19 + i, 6).setValue(dados.itens[i].quantidade || '');
        abaProposta.getRange(19 + i, 7).setValue(dados.itens[i].prazoEntrega || '');
        abaProposta.getRange(19 + i, 8).setValue(dados.itens[i].quantidadeEstoque || '');
        abaProposta.getRange(19 + i, 9).setValue(dados.itens[i].valorUnitario || '');
        abaProposta.getRange(19 + i, 10).setValue(dados.itens[i].desconto || '');
        abaProposta.getRange(19 + i, 11).setValue(dados.itens[i].valorUnitarioDesconto || '');
        abaProposta.getRange(19 + i, 12).setValue(dados.itens[i].valorTotalItem || '');
      }
      
      Logger.log("Dados dos itens gravados com sucesso.");
    } catch (error) {
      Logger.log("Erro ao gravar os dados: " + error.message);
    }
  } else {
    Logger.log("A aba especificada não foi encontrada.");
  }
}
