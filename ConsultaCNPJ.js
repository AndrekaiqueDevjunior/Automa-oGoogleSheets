// Função que cria um menu customizado
//function onOpen() {
 // var ui = SpreadsheetApp.getUi();
 // ui.createMenu('Consulta CNPJ')
   // .addItem('Buscar dados de CNPJ', 'abrirFormCNPJ')
   // .addToUi();
//}
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('FormCNPJ');
}



// Função que abre o formulário HTML
function abrirFormCNPJ() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('FormCNPJ')
      .setWidth(1000)
      .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Consulta CNPJ');
}

// Função para buscar e preencher os dados na planilha
// Função para preencher os dados na última linha disponível da planilha
/*function preencherDadosCNPJ(cnpj) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Chama a função que busca os dados do CNPJ
  var dados = buscarCNPJ(cnpj);

  // Verifica se houve erro na busca de dados
  if (dados[0].startsWith('Erro') || dados[0] === 'CNPJ inválido') {
    return dados[0];  // Retorna a mensagem de erro para o usuário
  }

  // Encontra a última linha disponível (após o último conteúdo)
  var ultimaLinha = sheet.getLastRow() + 1;

  // Insere os dados na última linha disponível
  sheet.getRange(ultimaLinha, 1, 1, dados.length).setValues([dados]);

  return dados;  // Retorna os dados para serem exibidos no HTML
}
 /* e */

// Função para buscar os dados de um CNPJ (usada no script anterior)
function buscarCNPJ(cnpj) {
  cnpj = cnpj.replace(/[^\d]+/g, '');
  if (cnpj.length !== 14) return ['CNPJ inválido'];
  
  var url = 'https://www.receitaws.com.br/v1/cnpj/' + cnpj;
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var result = JSON.parse(response.getContentText());
  
  if (result.status !== 'OK') return ['Erro: ' + result.message];
  
  return [
    result.nome,
    result.fantasia,
    result.uf,
    result.telefone,
    result.email,
    result.atividade_principal[0].text,
    result.situacao,
    result.logradouro,
    result.numero,
    result.bairro,
    result.municipio,
    result.capital_social,
    result.cnpj
  ];
}

// Função para buscar dados do CNPJ na célula
function ConsultarCNPJ(cnpj) {
  // Chama a função que busca os dados do CNPJ
  var dados = buscarCNPJ(cnpj);

  // Verifica se houve erro na busca de dados
  if (dados[0].startsWith('Erro') || dados[0] === 'CNPJ inválido') {
    return dados[0];  // Retorna a mensagem de erro
  }

  // Retorna os dados como uma matriz (array) para que sejam exibidos corretamente na planilha
  return [dados];  // Retorna os dados como um array de arrays
}

// Função para buscar os dados de um CNPJ (a mesma função anterior)
function buscarCNPJ(cnpj) {
  cnpj = cnpj.replace(/[^\d]+/g, '');
  if (cnpj.length !== 14) return ['CNPJ inválido'];
  
  var url = 'https://www.receitaws.com.br/v1/cnpj/' + cnpj;
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var result = JSON.parse(response.getContentText());
  
  if (result.status !== 'OK') return ['Erro: ' + result.message];
  
  return [
    result.nome, // [0]
    result.fantasia, // [1]
    result.uf, // [2]
    result.telefone, // [3]
    result.email, // [4]
    result.atividade_principal[0].text, // [5]
    result.situacao, // [6]
    result.logradouro, // [7]
    result.numero, // [8]
    result.bairro, // [9]
    result.municipio, // [10]
    result.capital_social, // [11]
    result.cnpj // [12]
  ];
}

// Função para buscar e preencher os dados na planilha "Proposta 2024"
function preencherDadosCNPJ(cnpj) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = planilha.getSheetByName("Proposta-2024"); // Seleciona a aba específica "Proposta 2024"
  
  if (!sheet) {
    return 'Erro: Aba "Proposta 2024" não encontrada.'; // Verifica se a aba existe
  }

  // Chama a função que busca os dados do CNPJ
  var dados = buscarCNPJ(cnpj);

  // Verifica se houve erro na busca de dados
  if (dados[0].startsWith('Erro') || dados[0] === 'CNPJ inválido') {
    return dados[0];  // Retorna a mensagem de erro para o usuário
  }

  // Preenche as células específicas na aba "Proposta 2024" com os dados
  sheet.getRange("A5:B5").merge(); // Mescla as células A5 e B5
  sheet.getRange("A5").setValue(dados[0]); // Nome da empresa

  sheet.getRange("C5:D5").merge(); // Mescla as células C5 e D5
  sheet.getRange("C5").setValue(dados[1]); // Nome fantasia

  sheet.getRange("J7").setValue(dados[2]); // UF
  sheet.getRange("K7").setValue(dados[3]); // Telefone

  sheet.getRange("J5:L5").merge(); // Mescla as células J5 até L5
  sheet.getRange("J5").setValue(dados[4]); // Email

  sheet.getRange("D7").setValue(dados[7]); // Logradouro
  sheet.getRange("E7").setValue(dados[8]); // Número
  sheet.getRange("G7").setValue(dados[9]); // Bairro
  sheet.getRange("I7").setValue(dados[10]); // Município

  sheet.getRange("E5:F5").merge(); // Mescla as células A5 e B5
  sheet.getRange("E5").setValue(dados[12]); //CNPJ  
  // Se houver outros campos para preencher, adicione-os conforme necessário
  // Exemplo:
  // sheet.getRange("M7").setValue(dados[5]); // Atividade principal
  // sheet.getRange("N7").setValue(dados[6]); // Situação
  // sheet.getRange("O7").setValue(dados[11]); // Capital social

  return dados;  // Retorna os dados para serem exibidos no HTML, se necessário
}
