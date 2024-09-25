function recuperarDadosDoHistorico(numeroProposta) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var abaProposta = spreadsheet.getSheetByName('Proposta-2024');
  var abaHistorico = spreadsheet.getSheetByName('Histórico');

  if (!abaProposta || !abaHistorico) {
    SpreadsheetApp.getUi().alert('Uma das abas especificadas não foi encontrada.');
    return;
  }

  // Procurar o número da proposta na aba Histórico
  var dadosHistorico = abaHistorico.getDataRange().getValues();
  var propostaEncontrada = false;

  for (var i = 0; i < dadosHistorico.length; i++) {
    if (dadosHistorico[i][2] == numeroProposta) { // Verifica se o número da proposta está correto

      // Definindo os intervalos a serem preenchidos
      var intervalos = [
        { range: 'A5', value: dadosHistorico[i][3] }, // Cliente
        { range: 'C5:D5', value: dadosHistorico[i][4] }, // Nome Fantasia
        { range: 'E5', value: dadosHistorico[i][5] }, // CNPJ
        { range: 'G5:I5', value: dadosHistorico[i][6] }, // Comprador
        { range: 'J5', value: dadosHistorico[i][7] }, // Email do Comprador
        { range: 'A7:B7', value: dadosHistorico[i][8] }, // Telefone
        { range: 'C7', value: dadosHistorico[i][9] }, // WhatsApp
        { range: 'D7', value: dadosHistorico[i][10] }, // Endereço
        { range: 'E7', value: dadosHistorico[i][11] }, // Número
        { range: 'F7', value: dadosHistorico[i][12] }, // Complemento
        { range: 'G7', value: dadosHistorico[i][13] }, // Bairro
        { range: 'I7', value: dadosHistorico[i][14] }, // Cidade
        { range: 'J7', value: dadosHistorico[i][15] }, // Estado
        { range: 'L7', value: dadosHistorico[i][16] }, // Telefone Geral
        { range: 'L7', value: dadosHistorico[i][17] }, // Email Geral
        { range: 'A9', value: dadosHistorico[i][18] }, // AD1
        { range: 'B9', value: dadosHistorico[i][19] }, // AD2
        { range: 'C9', value: dadosHistorico[i][20] }, // AD3
        { range: 'D9', value: dadosHistorico[i][21] }, // AD4
        { range: 'E9', value: dadosHistorico[i][22] }, // AD5
        { range: 'F9', value: dadosHistorico[i][23] }, // AD6
        { range: 'G9', value: dadosHistorico[i][24] }, // AD7
        { range: 'I9', value: dadosHistorico[i][25] }, // AD8
        { range: 'J9', value: dadosHistorico[i][26] }, // AD9
        { range: 'K9', value: dadosHistorico[i][27] }, // AD10
        { range: 'L9', value: dadosHistorico[i][28] }, // AD11
        { range: 'A12:D12', value: dadosHistorico[i][29] }, // Data Proposta
        { range: 'F12:L12', value: dadosHistorico[i][30] }, // Validade Proposta
        { range: 'A15:E15', value: dadosHistorico[i][31] }, // Condição de Pagamento
        { range: 'F15:G15', value: dadosHistorico[i][32] }, // Frete
        { range: 'I15:L15', value: dadosHistorico[i][33] }, // Transportadora
        { range: 'G46:L46', value: dadosHistorico[i][34] }, // Valor Total Proposta
        { range: 'G47:L47', value: dadosHistorico[i][35] }, // Impostos
        { range: 'G48', value: dadosHistorico[i][36] }, // Frete2
        { range: 'G49:L49', value: dadosHistorico[i][37] }, // Nota sobre Prazo de Entrega
        { range: 'G51:L51', value: dadosHistorico[i][38] }, // Vendedor
        { range: 'G52:L52', value: dadosHistorico[i][39] }, // Email Vendedor
        { range: 'G53:L53', value: dadosHistorico[i][40] }, // Telefone Vendedor
        { range: 'G54:L54', value: dadosHistorico[i][41] }, // WhatsApp Vendedor
      ];

      // Preencher as células com os dados da proposta
      intervalos.forEach(function (intervalo) {
        abaProposta.getRange(intervalo.range).setValue(intervalo.value);
      });

      // Preencher as colunas de item (A19:L44)
      for (var rowIndex = 19; rowIndex <= 44; rowIndex++) {
        var baseIndex = 42 + (rowIndex - 19) * 11; // Cálculo do índice da coluna no histórico

        abaProposta.getRange(rowIndex, 1).setValue(dadosHistorico[i][baseIndex]);       // Coluna A (Item)
        abaProposta.getRange(rowIndex, 2).setValue(dadosHistorico[i][baseIndex + 1]);   // Coluna B (Fabricante)
        abaProposta.getRange(rowIndex, 3).setValue(dadosHistorico[i][baseIndex + 2]);   // Coluna C (Código)
        abaProposta.getRange(rowIndex, 4).setValue(dadosHistorico[i][baseIndex + 3]);   // Coluna D (Descrição)
        abaProposta.getRange(rowIndex, 5).setValue(dadosHistorico[i][baseIndex + 4]);   // Coluna E (NCM)
        abaProposta.getRange(rowIndex, 6).setValue(dadosHistorico[i][baseIndex + 5]);   // Coluna F (QTde)
        abaProposta.getRange(rowIndex, 7).setValue(dadosHistorico[i][baseIndex + 6]);   // Coluna G (Prazo de Entrega)
        abaProposta.getRange(rowIndex, 8).setValue(dadosHistorico[i][baseIndex + 7]);   // Coluna I (Quantidade em Estoque)
        abaProposta.getRange(rowIndex, 9).setValue(dadosHistorico[i][baseIndex + 8]);   // Coluna J (Valor Unitário)
        abaProposta.getRange(rowIndex, 10).setValue(dadosHistorico[i][baseIndex + 9]);  // Coluna K (Desconto)
        abaProposta.getRange(rowIndex, 11).setValue(dadosHistorico[i][baseIndex + 10]); // Coluna L (Valor Unitário c/ Desconto)
      }

      propostaEncontrada = true; // Define que a proposta foi encontrada
      break; // Para sair do loop após encontrar a proposta
    }
  }

  // Alertar se a proposta não foi encontrada
  if (!propostaEncontrada) {
    SpreadsheetApp.getUi().alert('Número da proposta não encontrado no histórico.');
  }
}

function abrirRecuperarDados() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Recuperar Dados', 'Digite o número da proposta:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    var numeroProposta = response.getResponseText();
    recuperarDadosDoHistorico(numeroProposta);
  }
}
