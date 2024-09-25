function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Gera Proposta e CNPJ')
    .addItem('Buscar dados de CNPJ', 'abrirFormCNPJ')
    .addItem('Histórico de Consultas', 'abrirHistoricoConsultas')
    .addItem('Gerar Proposta 2024', 'enviarPdfParaDriveiexportarParaHistorico')
    .addItem('Abrir Consulta', 'abrirConsulta')
    .addItem('Recuperar Proposta', 'abrirRecuperarDados') // Adicione aqui
    .addToUi();

}

  function abrirConsulta() {
    var html = HtmlService.createHtmlOutputFromFile('ConsultaModal')
      .setWidth(1280)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Consulta de Propostas');
  }



  //function onOpen() {
  // var ui = SpreadsheetApp.getUi();
  // ui.createMenu('Formulário')
  //   .addItem('Abrir Formulário', 'abrirFormulario')
  //   .addToUi();
  //}
  function enviarPdfParaDriveiexportarParaHistorico() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var abaProposta = spreadsheet.getSheetByName('Proposta-2024');
    var abaHistorico = spreadsheet.getSheetByName('Histórico');

    if (!abaProposta || !abaHistorico) {
      SpreadsheetApp.getUi().alert('Uma das abas especificadas não foi encontrada.');
      return;
    }

    // Obter o número da proposta
    var hoje = new Date();
    var ano = hoje.getFullYear().toString().slice(-2); // Últimos 2 dígitos do ano
    var mes = (hoje.getMonth() + 1).toString().padStart(2, '0'); // Mês
    var dia = hoje.getDate().toString().padStart(2, '0'); // Dia
    var sequencia = (abaHistorico.getLastRow() + 1).toString().padStart(4, '0'); // Sequência
    var revisao = '0'; // Revisão sempre 0

    var numeroProposta = `${ano}${mes}${dia}${sequencia}-${revisao}`;
    abaProposta.getRange('I1').setValue(numeroProposta);
    var linkPDF = gerarPDF(abaProposta, numeroProposta);

    // Coletar dados da aba "Proposta Novus"
    var dados = {
      cliente: abaProposta.getRange('A5').getValue() + ' ' + abaProposta.getRange('B5').getValue() + ' ' +
        abaProposta.getRange('C5').getValue() + ' ' + abaProposta.getRange('D5').getValue(),
      nomeFantasia: abaProposta.getRange('C5:D5').getValue(),
      cnpj: abaProposta.getRange('E5').getValue() + ' ' + abaProposta.getRange('F5').getValue(),
      comprador: abaProposta.getRange('G5:I5').getValue() + ' ' + abaProposta.getRange('H5').getValue() + ' ' + abaProposta.getRange('I5').getValue(),

      emailComprador: abaProposta.getRange('J5').getValue() + ' ' + abaProposta.getRange('K5').getValue() + ' ' + abaProposta.getRange('L5').getValue(),

      telefone: abaProposta.getRange('A7:B7').getValue() + ' ' + abaProposta.getRange('B7').getValue(),

      whatsapp: abaProposta.getRange('C7').getValue(),

      endereco: abaProposta.getRange('D7').getValue(),

      numero: abaProposta.getRange('E7').getValue(),

      complemento: abaProposta.getRange('F7').getValue(),

      bairro: abaProposta.getRange('G7').getValue(),

      cidade: abaProposta.getRange('I7').getValue(),
      estado: abaProposta.getRange('J7').getValue(),
      telefoneGeral: abaProposta.getRange('L7').getValue(),
      emailGeral: abaProposta.getRange('L7').getValue(),
      ad1: abaProposta.getRange('A9').getValue(),
      ad2: abaProposta.getRange('B9').getValue(),
      ad3: abaProposta.getRange('C9').getValue(),
      ad4: abaProposta.getRange('D9').getValue(),
      ad5: abaProposta.getRange('E9').getValue(),
      ad6: abaProposta.getRange('F9').getValue(),
      ad7: abaProposta.getRange('G9').getValue(),
      ad8: abaProposta.getRange('I9').getValue(),
      ad9: abaProposta.getRange('J9').getValue(),
      ad10: abaProposta.getRange('K9').getValue(),
      ad11: abaProposta.getRange('L9').getValue(),
      dataProposta: formatarData(abaProposta.getRange('A12:D12').getValue()) + ' ' + formatarData(abaProposta.getRange('B12').getValue()) + ' ' + formatarData(abaProposta.getRange('C12').getValue()) + ' ' + formatarData(abaProposta.getRange('D12').getValue()),
      validadeProposta: formatarData(abaProposta.getRange('F12:L12').getValue()), // Corrigido aqui
      condicaoPagamento: abaProposta.getRange('A15:E15').getValue(),
      frete: abaProposta.getRange('F15:G15').getValue(),
      transportadora: abaProposta.getRange('I15:L15').getValue(),
      valorTotalProposta: abaProposta.getRange('G46:L46').getValue(),
      impostos: abaProposta.getRange('G47:L47').getValues(),
      frete2: abaProposta.getRange('G48').getValue(),
      notasobrePrazoEntrega: abaProposta.getRange('G49:L49').getValue(),


      vendedor: abaProposta.getRange('G51:L51').getValue() + ' ' + abaProposta.getRange('H49').getValue() + ' ' + abaProposta.getRange('I49').getValue() + ' ' + abaProposta.getRange('J49').getValue() + ' ' + abaProposta.getRange('K49').getValue() + ' ' + abaProposta.getRange('L49').getValue(),

      emailVendedor: abaProposta.getRange('G52:L52').getValue() + ' ' + abaProposta.getRange('H50').getValue() + ' ' + abaProposta.getRange('I50').getValue() + ' ' + abaProposta.getRange('J50').getValue() + ' ' + abaProposta.getRange('K50').getValue() + ' ' + abaProposta.getRange('L50').getValue(),

      telefoneVendedor: abaProposta.getRange('G53:L53').getValue() + ' ' + abaProposta.getRange('H51').getValue() + ' ' + abaProposta.getRange('I51').getValue() + ' ' + abaProposta.getRange('J51').getValue() + ' ' + abaProposta.getRange('K51').getValue() + ' ' + abaProposta.getRange('L51').getValue(),



      whatsappVendedor: abaProposta.getRange('G54:L54').getValue() + ' ' + abaProposta.getRange('H52').getValue() + ' ' + abaProposta.getRange('I52').getValue() + ' ' + abaProposta.getRange('J52').getValue() + ' ' + abaProposta.getRange('K52').getValue() + ' ' + abaProposta.getRange('L52').getValue(),

      item: abaProposta.getRange('A19').getValue(),
      fabricante: abaProposta.getRange('B19').getValue(),
      codigo: abaProposta.getRange('C19').getValue(),
      descricao: abaProposta.getRange('D19').getValue(),
      ncm: abaProposta.getRange('E19').getValue(),
      quantidade: abaProposta.getRange('F19').getValue(),
      prazoEntrega: abaProposta.getRange('G19').getValue(),
      quantidadeEstoque: abaProposta.getRange('H19').getValue(),
      valorUnitario: abaProposta.getRange('I19').getValue(),
      desconto: abaProposta.getRange('J19').getValue(),
      valorUnitarioDesconto: abaProposta.getRange('K19').getValue(),
      valorTotalItem: abaProposta.getRange('L19').getValue(),
      numeroProposta: numeroProposta
    };

    // Obter dados dos itens
    var intervaloItens = abaProposta.getRange('A19:L' + abaProposta.getLastRow());
    var itens = intervaloItens.getValues();

    // Preencher a aba Histórico com os dados
    var linhaHistorico = abaHistorico.getLastRow() + 1; // Próxima linha disponível

    abaHistorico.getRange('A' + linhaHistorico).setValue(new Date()); // Data e hora da execução
    abaHistorico.getRange('B' + linhaHistorico).setValue(linkPDF);
    abaHistorico.getRange('C' + linhaHistorico).setValue(numeroProposta); // Número da proposta gerada
    abaHistorico.getRange('D' + linhaHistorico).setValue(dados.cliente);
    abaHistorico.getRange('E' + linhaHistorico).setValue(dados.nomeFantasia);
    abaHistorico.getRange('F' + linhaHistorico).setValue(dados.cnpj);
    abaHistorico.getRange('G' + linhaHistorico).setValue(dados.comprador);
    abaHistorico.getRange('H' + linhaHistorico).setValue(dados.emailComprador);
    abaHistorico.getRange('I' + linhaHistorico).setValue(dados.telefone);
    abaHistorico.getRange('J' + linhaHistorico).setValue(dados.whatsapp);
    abaHistorico.getRange('K' + linhaHistorico).setValue(dados.endereco);
    abaHistorico.getRange('L' + linhaHistorico).setValue(dados.numero);
    abaHistorico.getRange('M' + linhaHistorico).setValue(dados.complemento);
    abaHistorico.getRange('N' + linhaHistorico).setValue(dados.bairro);
    abaHistorico.getRange('O' + linhaHistorico).setValue(dados.cidade);
    abaHistorico.getRange('P' + linhaHistorico).setValue(dados.estado);
    abaHistorico.getRange('Q' + linhaHistorico).setValue(dados.telefoneGeral);
    abaHistorico.getRange('R' + linhaHistorico).setValue(dados.emailGeral);
    abaHistorico.getRange('S' + linhaHistorico).setValue(dados.ad1);
    abaHistorico.getRange('T' + linhaHistorico).setValue(dados.ad2);
    abaHistorico.getRange('U' + linhaHistorico).setValue(dados.ad3);
    abaHistorico.getRange('V' + linhaHistorico).setValue(dados.ad4);
    abaHistorico.getRange('W' + linhaHistorico).setValue(dados.ad5);
    abaHistorico.getRange('X' + linhaHistorico).setValue(dados.ad6);
    abaHistorico.getRange('Y' + linhaHistorico).setValue(dados.ad7);
    abaHistorico.getRange('Z' + linhaHistorico).setValue(dados.ad8);
    abaHistorico.getRange('AA' + linhaHistorico).setValue(dados.ad9);
    abaHistorico.getRange('AB' + linhaHistorico).setValue(dados.ad10);
    abaHistorico.getRange('AC' + linhaHistorico).setValue(dados.ad11);
    abaHistorico.getRange('AD' + linhaHistorico).setValue(dados.dataProposta);
    abaHistorico.getRange('AE' + linhaHistorico).setValue(dados.validadeProposta);
    abaHistorico.getRange('AF' + linhaHistorico).setValue(dados.condicaoPagamento);
    abaHistorico.getRange('AG' + linhaHistorico).setValue(dados.frete);
    abaHistorico.getRange('AH' + linhaHistorico).setValue(dados.transportadora);
    abaHistorico.getRange('AI' + linhaHistorico).setValue(dados.valorTotalProposta);
    abaHistorico.getRange('AJ' + linhaHistorico).setValue(dados.impostos);
    abaHistorico.getRange('AK' + linhaHistorico).setValue(dados.frete2);
    abaHistorico.getRange('AL' + linhaHistorico).setValue(dados.notasobrePrazoEntrega);

    abaHistorico.getRange('AM' + linhaHistorico).setValue(dados.vendedor);
    abaHistorico.getRange('AN' + linhaHistorico).setValue(dados.emailVendedor);
    abaHistorico.getRange('AO' + linhaHistorico).setValue(dados.telefoneVendedor);
    abaHistorico.getRange('AP' + linhaHistorico).setValue(dados.whatsappVendedor);

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
    var ultimaLinhaHistorico = abaHistorico.getLastRow();

    // Determinar a próxima linha disponível na aba "Histórico"
    var linhaInicioHistorico = ultimaLinhaHistorico + 0;

    // Determinar a coluna inicial e número de colunas usadas
    var colunaInicial = 43; // Coluna AQ corresponde à coluna 43
    var numeroDeColunas = linhaConcatenada.length;

    // Inserir a linha concatenada na aba "Histórico" a partir da próxima linha disponível e coluna AQ
    var intervaloDestino = abaHistorico.getRange(linhaInicioHistorico, colunaInicial, 1, numeroDeColunas);
    intervaloDestino.setValues([linhaConcatenada]);


    mostrarMensagemHTML(numeroProposta, linkPDF);
  }

  function mostrarMensagemHTML(numeroProposta, linkPDF) {
    return HtmlService.createHtmlOutput('<p>A proposta gerada com sucesso!</p>' +
      '<p>Número da proposta: ' + numeroProposta + '</p>' +
      '<p><a href="' + linkPDF + '" target="_blank">Clique aqui para visualizar o PDF</a></p>' +
      '<p><button onclick="google.script.host.close()">Fechar</button></p>')
      .setWidth(400)
      .setHeight(300);
  }

  function gerarPDF(abaProposta, numeroProposta) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(abaProposta); // Usando abaProposta diretamente

  if (!sheet) {
    throw new Error('Aba não encontrada: ' + abaProposta);
  }

  // Verifica o valor da célula I1
  var valorI1 = sheet.getRange('I1').getValue();
  if (valorI1 !== 'gerarProposta' && valorI1 !== 'nova proposta') {
    throw new Error('A geração do PDF só é permitida se a célula I1 contiver "gerarProposta" ou "nova proposta".');
  }

  // Configura as opções de exportação para o PDF
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf' +
    '&gid=' + sheet.getSheetId() +
    '&portrait=true' +  // Orientação da página (paisagem ou retrato)
    '&size=A4' +          // Tamanho do papel
    '&fitw=true' +        // Ajustar largura
    '&gridlines=false' +  // Ocultar linhas de grade
    '&printtitle=false' + // Ocultar títulos
    '&sheetnames=false' + // Ocultar nomes das abas
    '&pagenumbers=false' + // Ocultar números das páginas
    '&horizontal_alignment=CENTER' + // Alinhamento horizontal
    '&vertical_alignment=TOP'; // Alinhamento vertical

  // Faz a requisição para exportar o PDF
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });

  // Obtém o conteúdo do PDF
  var pdfBlob = response.getBlob().setName(`Proposta_${numeroProposta}.pdf`);

  // Cria o arquivo PDF na pasta desejada
  var folder = DriveApp.getFolderById("10ZbsX--0llWDx4grgBANet1xzqAEICNF"); // ID da pasta no Google Drive
  var arquivoPDF = folder.createFile(pdfBlob);

  var destinatario = 'anderekaidellisola@gmail.com';
  var assunto = 'Proposta PDF ' + numeroProposta; // Inclui o número da proposta no assunto
  var corpoEmail = 'Olá,\n\nSegue em anexo o PDF da proposta número ' + numeroProposta + '.\n\nAtenciosamente,\nSua Equipe';

  MailApp.sendEmail({
    to: destinatario,
    subject: assunto,
    body: corpoEmail,
    attachments: [pdfBlob]
  });

  return arquivoPDF.getUrl(); // Retorna o link do arquivo PDF
}


  function formatarData(data) {
    if (data instanceof Date) {
      return Utilities.formatDate(data, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    return '';
  }

  function gerarPDFEQRCode(abaProposta, numeroProposta) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(abaProposta);

    if (!sheet) {
      throw new Error('Aba não encontrada: ' + abaProposta);
    }

    // Configura as opções de exportação para o PDF
    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf' +
      '&gid=' + sheet.getSheetId() +
      '&portrait=true' +
      '&size=A4' +
      '&fitw=true' +
      '&gridlines=false' +
      '&printtitle=false' +
      '&sheetnames=false' +
      '&pagenumbers=false' +
      '&horizontal_alignment=CENTER' +
      '&vertical_alignment=TOP';

    // Faz a requisição para exportar o PDF
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });

    // Obtém o conteúdo do PDF
    var pdfBlob = response.getBlob().setName(`Proposta_${numeroProposta}.pdf`);

    // Cria o arquivo PDF na pasta desejada
    var folder = DriveApp.getFolderById("10ZbsX--0llWDx4grgBANet1xzqAEICNF");
    var arquivoPDF = folder.createFile(pdfBlob);

    // Gera o link do PDF
    var linkPDF = arquivoPDF.getUrl();

    // Verifica se o linkPDF está correto
    if (!linkPDF) {
      throw new Error('Erro ao gerar link do PDF.');
    }

    // Gera o QR Code
    var qrCodeUrl = gerarQRCode(linkPDF);

    // Gera o QR Code usando a API
    var qrCodeBlob = UrlFetchApp.fetch(qrCodeUrl).getBlob().setName(`QRCode_Proposta_${numeroProposta}.png`);
    folder.createFile(qrCodeBlob);

    return { pdfUrl: linkPDF, qrCodeUrl: qrCodeUrl };
  }

  function gerarQRCode(link) {
    // Encoda o link para garantir que não há caracteres especiais
    var encodedLink = encodeURIComponent(link);
    var qrCodeApiUrl = `https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=${encodedLink}`;
    return qrCodeApiUrl;
  }