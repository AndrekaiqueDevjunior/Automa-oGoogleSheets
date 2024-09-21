function backupPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataAtual = new Date();
  var dataFormatada = Utilities.formatDate(dataAtual, 'America/Sao_Paulo', 'dd/MM/yyyy');
  var nomeBackup = ss.getName() + "_Backup_" + dataFormatada.replace(/\//g, '-');

  // Copia a planilha para o Google Drive
  var novaPlanilha = ss.copy(nomeBackup);

  // Define a pasta de destino
  var pastaBackupId = '1CrrogbvpcZyskF4e-rcVpJpV28_06f7o';
  var pastaBackup = DriveApp.getFolderById(pastaBackupId);
  var arquivoBackup = DriveApp.getFileById(novaPlanilha.getId());

  // Move o backup para a pasta desejada
  pastaBackup.addFile(arquivoBackup);
  DriveApp.getRootFolder().removeFile(arquivoBackup);

  // Envia uma notificação de backup
  MailApp.sendEmail({
    to: 'andrekaidellisola@gmail.com',
    subject: 'Backup diário realizado',
    body: 'O backup da planilha foi realizado com sucesso em ' + dataFormatada + '.'
  });
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Meu Menu')
      .addItem('Gerar PDF e QR Code', 'abrirModalComProposta')
      .addToUi();
}

function abrirModalComProposta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaProposta = "Proposta-2024";
  var sheet = ss.getSheetByName(abaProposta);

  if (!sheet) {
    throw new Error('Aba não encontrada: ' + abaProposta);
  }

  var numeroProposta = sheet.getRange("I1").getValue();
  var result = gerarPDFEQRCode(abaProposta); // Chame a função de geração

  var htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            text-align: center;
            padding: 20px;
          }
          h1 {
            color: #333;
          }
          h2 {
            color: #555;
          }
          a {
            display: inline-block;
            margin-top: 10px;
            padding: 10px 20px;
            background-color: #4CAF50;
            color: white;
            text-decoration: none;
            border-radius: 5px;
          }
          a:hover {
            background-color: #45a049;
          }
          img {
            margin-top: 20px;
          }
        </style>
      </head>
      <body>
        <h1>Número da Proposta: ${numeroProposta}</h1>
        <a href="${result.pdfUrl}" target="_blank">Baixar Proposta (PDF)</a>
        <h2>QR Code:</h2>
        <img src="${result.qrCodeUrl}" alt="QR Code">
      </body>
    </html>`;

  abrirModal(htmlContent); // Abra o modal com o conteúdo HTML
}

function abrirModal(htmlContent) {
  var html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Proposta Gerada');
}


function backupPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Obtém a planilha atual
  var dataAtual = new Date(); // Data e hora atual
  var dataFormatada = Utilities.formatDate(dataAtual, 'America/Sao_Paulo', 'dd/MM/yyyy'); // Formata a data no padrão brasileiro
  var nomeBackup = ss.getName() + "_Backup_" + dataFormatada.replace(/\//g, '-'); // Nome do backup (substitui "/" por "-" para evitar problemas no nome do arquivo)

  // Copia a planilha para o Google Drive
  var novaPlanilha = ss.copy(nomeBackup);

  // Define a pasta de destino
  var pastaBackupId = '1CrrogbvpcZyskF4e-rcVpJpV28_06f7o'; // Substitua pelo ID da pasta de backup no Google Drive
  var pastaBackup = DriveApp.getFolderById(pastaBackupId);

  // Move o backup para a pasta desejada
  var arquivoBackup = DriveApp.getFileById(novaPlanilha.getId());
  pastaBackup.addFile(arquivoBackup);

  // Remove o arquivo da pasta raiz do Drive
  DriveApp.getRootFolder().removeFile(arquivoBackup);

  // Envia uma notificação de backup
  MailApp.sendEmail({
    to: 'andrekaidellisola@gmail.com', // Seu email para notificação
    subject: 'Backup diário realizado',
    body: 'O backup da planilha foi realizado com sucesso em ' + dataFormatada + '.'
  });
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Meu Menu')
      .addItem('Gerar PDF e QR Code', 'gerarPDFEQRCode')
      .addToUi();
}

function gerarPDFEQRCode(numeroProposta) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaProposta = "Proposta-2024"; // Definindo a aba aqui
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

  try {
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

    // Cria o conteúdo HTML com a biblioteca QR Code
    var htmlContent = `<!DOCTYPE html>
      <html>
        <head>
          <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
          <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.qrcode/1.0/jquery.qrcode.min.js"></script>
          <style>
            body { font-family: Arial, sans-serif; }
            h1 { color: #333; }
            h2 { color: #666; }
            a { color: #007BFF; text-decoration: none; }
            a:hover { text-decoration: underline; }
          </style>
        </head>
        <body>
          <h1>Link da Proposta</h1>
          <a href="${linkPDF}" target="_blank">${linkPDF}</a>
          <h2>QR Code:</h2>
          <div id="qrcode"></div>
          <script>
            $(document).ready(function() {
              $('#qrcode').qrcode('${linkPDF}');
            });
          </script>
        </body>
      </html>`;

    // Chama a função para abrir o modal
    abrirModal(htmlContent);

    return { pdfUrl: linkPDF };
  } catch (e) {
    Logger.log('Erro: ' + e.message);
    throw new Error('Ocorreu um erro ao gerar o PDF ou o QR Code: ' + e.message);
  }
}

function abrirModal(htmlContent) {
  var html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(1240)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Proposta Gerada');
}

