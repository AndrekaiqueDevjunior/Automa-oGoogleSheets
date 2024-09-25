function abrirDialogoNumeroProposta() {
  var html = `
    <style>
      body {
        font-family: Arial, sans-serif;
        padding: 20px;
      }
      h2 {
        color: #4CAF50;
      }
      input[type="number"] {
        width: 100%;
        padding: 10px;
        margin: 10px 0;
        border: 1px solid #ccc;
        border-radius: 4px;
      }
      button {
        background-color: #4CAF50;
        color: white;
        padding: 10px 15px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
      }
      button:hover {
        background-color: #45a049;
      }
    </style>
    <body>
      <h2>Insira o número da proposta</h2>
      <input type="number" id="numeroProposta" placeholder="Número da proposta" />
      <button onclick="enviarNumeroProposta()">Buscar</button>

      <script>
        function enviarNumeroProposta() {
          var numeroProposta = document.getElementById('numeroProposta').value;
          google.script.run.recuperarDadosParaProposta(numeroProposta);
          google.script.host.close();
        }
      </script>
    </body>
  `;
  
  var ui = HtmlService.createHtmlOutput(html).setWidth(300).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Buscar Proposta');
}

function recuperarDadosParaProposta(numeroProposta) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var abaProposta = spreadsheet.getSheetByName('Proposta-2024');
  var abaHistorico = spreadsheet.getSheetByName('Histórico');

  if (!abaProposta || !abaHistorico) {
    SpreadsheetApp.getUi().alert('Uma das abas especificadas não foi encontrada.');
    return;
  }

  var intervaloHistorico = abaHistorico.getRange(1, 1, abaHistorico.getLastRow(), abaHistorico.getLastColumn());
  var historicoValores = intervaloHistorico.getValues();
  
  // Procurar o número da proposta no histórico
  var linhaHistorico = -1;
  for (var i = 0; i < historicoValores.length; i++) {
    if (historicoValores[i][2] == numeroProposta) { // A coluna 3 contém o número da proposta
      linhaHistorico = i + 1; // Armazenar a linha do histórico encontrada
      break;
    }
  }
  
  if (linhaHistorico == -1) {
    SpreadsheetApp.getUi().alert('Número da proposta não encontrado no histórico.');
    return;
  }

  // Mostrar loading antes de começar a recuperação dos dados
  mostrarMensagemCarregando();

  // Recuperar os dados para as respectivas células na aba Proposta-2024
  abaProposta.getRange('A5').setValue(historicoValores[linhaHistorico - 1][3]);  // Cliente
  abaProposta.getRange('C5').setValue(historicoValores[linhaHistorico - 1][4]);  // Nome Fantasia
  abaProposta.getRange('E5').setValue(historicoValores[linhaHistorico - 1][5]);  // CNPJ
  abaProposta.getRange('G5').setValue(historicoValores[linhaHistorico - 1][6]);  // Comprador
  abaProposta.getRange('J5').setValue(historicoValores[linhaHistorico - 1][7]);  // Email Comprador
  abaProposta.getRange('A7').setValue(historicoValores[linhaHistorico - 1][8]);  // Telefone
  abaProposta.getRange('C7').setValue(historicoValores[linhaHistorico - 1][9]);  // Whatsapp
  abaProposta.getRange('D7').setValue(historicoValores[linhaHistorico - 1][10]); // Endereço
  abaProposta.getRange('E7').setValue(historicoValores[linhaHistorico - 1][11]); // Número
  abaProposta.getRange('F7').setValue(historicoValores[linhaHistorico - 1][12]); // Complemento
  abaProposta.getRange('G7').setValue(historicoValores[linhaHistorico - 1][13]); // Bairro
  abaProposta.getRange('I7').setValue(historicoValores[linhaHistorico - 1][14]); // Cidade
  abaProposta.getRange('J7').setValue(historicoValores[linhaHistorico - 1][15]); // Estado
  abaProposta.getRange('L7').setValue(historicoValores[linhaHistorico - 1][16]); // Telefone Geral
  abaProposta.getRange('L7').setValue(historicoValores[linhaHistorico - 1][17]); // Email Geral

  // Preencher os dados adicionais (ad1 até ad11)
  abaProposta.getRange('A9').setValue(historicoValores[linhaHistorico - 1][18]);
  abaProposta.getRange('B9').setValue(historicoValores[linhaHistorico - 1][19]);
  abaProposta.getRange('C9').setValue(historicoValores[linhaHistorico - 1][20]);
  abaProposta.getRange('D9').setValue(historicoValores[linhaHistorico - 1][21]);
  abaProposta.getRange('E9').setValue(historicoValores[linhaHistorico - 1][22]);
  abaProposta.getRange('F9').setValue(historicoValores[linhaHistorico - 1][23]);
  abaProposta.getRange('G9').setValue(historicoValores[linhaHistorico - 1][24]);
  abaProposta.getRange('I9').setValue(historicoValores[linhaHistorico - 1][25]);
  abaProposta.getRange('J9').setValue(historicoValores[linhaHistorico - 1][26]);
  abaProposta.getRange('K9').setValue(historicoValores[linhaHistorico - 1][27]);
  abaProposta.getRange('L9').setValue(historicoValores[linhaHistorico - 1][28]);

  // Preencher os itens da proposta (evitando duplicação)
  var linhaInicial = 19;
  var colunaInicial = 1;
  var ultimaColuna = 12;

  // Capturar a linha com os itens
  var intervaloItens = abaHistorico.getRange(linhaHistorico, 43, 1, ultimaColuna);
  var itensProposta = intervaloItens.getValues();

  // Limpar a área onde os itens serão inseridos para evitar sobreposição
  abaProposta.getRange('A19:L44' + (abaProposta.getLastRow())).clearContent();

  // Inserir itens na aba Proposta-2024
  abaProposta.getRange(linhaInicial, colunaInicial, 1, itensProposta[0].length).setValues(itensProposta);

  // Esconder a mensagem de carregamento quando o processo for concluído
  SpreadsheetApp.getUi().alert('Dados recuperados com sucesso!');
}

function mostrarMensagemCarregando() {
  var html = `
    <style>
      body {
        font-family: Arial, sans-serif;
        margin: 20px;
        padding: 20px;
        border: 1px solid #ccc;
        border-radius: 8px;
        background-color: #f9f9f9;
      }
      h2 {
        color: #4CAF50;
      }
      p {
        font-size: 16px;
        text-align: center;
      }
      .loader {
        border: 16px solid #f3f3f3;
        border-radius: 50%;
        border-top: 16px solid #3498db;
        width: 120px;
        height: 120px;
        animation: spin 2s linear infinite;
        margin: 20px auto;
      }
      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }
    </style>
    <body>
      <h2>Aguarde...</h2>
      <p>Carregando dados da proposta</p>
      <div class="loader"></div>
    </body>
  `;
  
  var ui = HtmlService.createHtmlOutput(html).setWidth(300).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Carregando');
}
