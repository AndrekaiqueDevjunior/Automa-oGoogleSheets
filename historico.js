// Função que abre a aba de histórico de consultas
function abrirHistoricoConsultas() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('HistoricoConsultas')
      .setWidth(1000)
      .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Histórico de Consultas');
}

function getHistoricoConsultas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ConsultarCNPJ'); // Substitua pelo nome da sua aba de histórico
  if (!sheet) {
    return []; // Retorna um array vazio se a aba não existir
  }
  
  var dados = sheet.getDataRange().getValues(); // Pega todos os dados da planilha
  var historico = [];

  // Ignora a primeira linha se for cabeçalho
  for (var i = 1; i < dados.length; i++) {
    historico.push(dados[i]);
  }

  return historico; // Retorna o histórico para o HTML
}

