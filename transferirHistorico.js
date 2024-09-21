function transferirDadosPropostaParaHistoricoHorizontal() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var abaProposta = spreadsheet.getSheetByName('Proposta-2024');
  var abaHistorico = spreadsheet.getSheetByName('Histórico');
  
  if (!abaProposta || !abaHistorico) {
    SpreadsheetApp.getUi().alert('Uma das abas especificadas não foi encontrada.');
    return;
  }
  
  // Obter os dados do intervalo A19:L44 na aba "Proposta-2024"
  var intervaloProposta = abaProposta.getRange('A19:L44');
  var dadosProposta = intervaloProposta.getValues(); // Array 2D com os dados

  // Concatenar todos os valores em uma única linha
  var linhaConcatenada = [];
  
  for (var i = 0; i < dadosProposta.length; i++) {
    for (var j = 0; j < dadosProposta[i].length; j++) {
      linhaConcatenada.push(dadosProposta[i][j]);
    }
  }

  // Obter a última linha preenchida na aba "Histórico"
  //var ultimaLinhaHistorico = abaHistorico.getLastRow();
  
  // Determinar a próxima linha disponível na aba "Histórico"
  //var linhaInicioHistorico = ultimaLinhaHistorico + 1;

  // Determinar a coluna inicial e número de colunas usadas
  var colunaInicial = 43; // Coluna AQ corresponde à coluna 43
  var numeroDeColunas = linhaConcatenada.length;

  // Inserir a linha concatenada na aba "Histórico" a partir da próxima linha disponível e coluna AQ
  var intervaloDestino = abaHistorico.getRange(colunaInicial, 1, numeroDeColunas);
  intervaloDestino.setValues([linhaConcatenada]);

  SpreadsheetApp.getUi().alert('Dados transferidos com sucesso da aba "Proposta-2024" para a aba "Histórico" de forma horizontal, iniciando na coluna AQ.');
}
