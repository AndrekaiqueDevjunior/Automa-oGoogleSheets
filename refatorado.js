function enviarPdfParaDriveiexportarParaHistorico() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var abaProposta = spreadsheet.getSheetByName('Proposta-2024');
  var abaHistorico = spreadsheet.getSheetByName('Histórico');

  if (!abaProposta || !abaHistorico) {
    SpreadsheetApp.getUi().alert('Uma das abas especificadas não foi encontrada.');
    return;
  }

  // Gerar o número da proposta
  var numeroProposta = gerarNumeroProposta(abaHistorico);
  abaProposta.getRange('I1').setValue(numeroProposta);
  var linkPDF = gerarPDF(abaProposta, numeroProposta);

  // Coletar dados da aba "Proposta-2024"
  var dados = coletarDados(abaProposta, numeroProposta);

  // Preencher a aba "Histórico" com os dados
  preencherHistorico(abaHistorico, dados, linkPDF, numeroProposta);

  SpreadsheetApp.getUi().alert('Dados transferidos com sucesso da aba "Proposta-2024" para a aba "Histórico" de forma horizontal, iniciando na coluna AQ.');
}

function gerarNumeroProposta(abaHistorico) {
  var hoje = new Date();
  var ano = hoje.getFullYear().toString().slice(-2);
  var mes = (hoje.getMonth() + 1).toString().padStart(2, '0');
  var dia = hoje.getDate().toString().padStart(2, '0');
  var sequencia = (abaHistorico.getLastRow() + 1).toString().padStart(4, '0');
  var revisao = '0';
  return `${ano}${mes}${dia}${sequencia}-${revisao}`;
}

function coletarDados(abaProposta, numeroProposta) {
  var dados = {};
  var ranges = {
    cliente: 'A5:D5',
    nomeFantasia: 'C5:D5',
    cnpj: 'E5:F5',
    comprador: 'G5:I5',
    emailComprador: 'J5:L5',
    telefone: 'A7:B7',
    whatsapp: 'C7',
    endereco: 'D7',
    numero: 'E7',
    complemento: 'F7',
    bairro: 'G7',
    cidade: 'I7',
    estado: 'J7',
    telefoneGeral: 'L7',
    emailGeral: 'L7',
    ad1: 'A9',
    ad2: 'B9',
    ad3: 'C9',
    ad4: 'D9',
    ad5: 'E9',
    ad6: 'F9',
    ad7: 'G9',
    ad8: 'I9',
    ad9: 'J9',
    ad10: 'K9',
    ad11: 'L9',
    dataProposta: 'A12:D12',
    validadeProposta: 'F12:L12',
    condicaoPagamento: 'A15:E15',
    frete: 'F15:G15',
    transportadora: 'I15:L15',
    valorTotalProposta: 'G46:L46',
    impostos: 'G47:L47',
    frete2: 'G48',
    notasobrePrazoEntrega: 'G49:L49',
    vendedor: 'G51:L51',
    emailVendedor: 'G52:L52',
    telefoneVendedor: 'G53:L53',
    whatsappVendedor: 'G54:L54'
  };

  for (var key in ranges) {
    var range = abaProposta.getRange(ranges[key]);
    dados[key] = key.includes('data') ? formatarData(range.getValue()) : range.getValue();
  }

  dados.numeroProposta = numeroProposta;

  // Coletar dados dos itens
  var intervaloItens = abaProposta.getRange('A19:L' + abaProposta.getLastRow());
  var itens = intervaloItens.getValues();
  dados.itensFiltrados = itens.filter((item, index) => index < 27); // Limita a 27 linhas

  return dados;
}

function preencherHistorico(abaHistorico, dados, linkPDF, numeroProposta) {
  var linhaHistorico = abaHistorico.getLastRow() + 1;

  abaHistorico.getRange('A' + linhaHistorico).setValue(new Date());
  abaHistorico.getRange('B' + linhaHistorico).setValue(linkPDF);
  abaHistorico.getRange('C' + linhaHistorico).setValue(numeroProposta);

  var colunas = ['cliente', 'nomeFantasia', 'cnpj', 'comprador', 'emailComprador', 'telefone', 'whatsapp', 'endereco', 'numero', 'complemento', 'bairro', 'cidade', 'estado', 'telefoneGeral', 'emailGeral', 'ad1', 'ad2', 'ad3', 'ad4', 'ad5', 'ad6', 'ad7', 'ad8', 'ad9', 'ad10', 'ad11', 'dataProposta', 'validadeProposta', 'condicaoPagamento', 'frete', 'transportadora', 'valorTotalProposta', 'impostos', 'frete2', 'notasobrePrazoEntrega', 'vendedor', 'emailVendedor', 'telefoneVendedor', 'whatsappVendedor'];

  colunas.forEach((coluna, index) => {
    abaHistorico.getRange(linhaHistorico, index + 4).setValue(dados[coluna]); // Começa em D
  });

  var linhaConcatenada = dados.itensFiltrados.flat(); // Concatenar os dados dos itens
  abaHistorico.getRange(linhaHistorico, 43, 1, linhaConcatenada.length).setValues([linhaConcatenada]); // Coluna AQ
}

function gerarPDF(abaProposta, numeroProposta) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Proposta-2024");

  if (!sheet) {
    throw new Error('Aba não encontrada: ' + abaProposta);
  }

  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf' +
            '&gid=' + sheet.getSheetId() +
            '&portrait=true&size=A4&fitw=true&gridlines=false&printtitle=false' +
            '&sheetnames=false&pagenumbers=false&horizontal_alignment=CENTER&vertical_alignment=TOP';

  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });

  var pdfBlob = response.getBlob().setName(`Proposta_${numeroProposta}.pdf`);
  var folder = DriveApp.getFolderById("10ZbsX--0llWDx4grgBANet1xzqAEICNF");
  var arquivoPDF = folder.createFile(pdfBlob);

  MailApp.sendEmail({
    to: 'andrekaidellisola@gmail.com',
    subject: 'Proposta PDF ' + numeroProposta,
    body: 'Olá,\n\nSegue em anexo o PDF da proposta número ' + numeroProposta + '.\n\nAtenciosamente,\nSua Equipe',
    attachments: [pdfBlob]
  });

  return arquivoPDF.getUrl();
}

function formatarData(data) {
  if (data instanceof Date) {
    return Utilities.formatDate(data, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return '';
}
