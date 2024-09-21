// Função que cria um menu customizado
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Consulta CNPJ')
    .addItem('Buscar dados de CNPJ', 'abrirFormCNPJ')
    .addToUi();
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
function preencherDadosCNPJ(cnpj) {
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
    result.capital_social
  ];
}
