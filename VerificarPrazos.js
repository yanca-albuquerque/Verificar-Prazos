function verificarPrazosEVouchers() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abas = planilha.getSheets();
  var emailsDestino = "email"; 
  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0); 
  var mensagens = [];

  abas.forEach(function(aba) {
    var dados = aba.getDataRange().getValues();
    var cabecalho = dados[0]; 
    var indiceData = cabecalho.indexOf("Validade");
    var indiceStatus = cabecalho.indexOf("Status");

    if (indiceData !== -1 && indiceStatus !== -1) {
      for (var i = 1; i < dados.length; i++) { 
        var dataVencimentoRaw = dados[i][indiceData];
        var status = dados[i][indiceStatus];
        
        if (dataVencimentoRaw instanceof Date && !isNaN(dataVencimentoRaw) && status !== "Entregue") {
          var dataVencimento = new Date(dataVencimentoRaw);
          dataVencimento.setHours(0, 0, 0, 0); 

          var diferencaDias = Math.floor((dataVencimento - hoje) / (1000 * 60 * 60 * 24));

          if (diferencaDias > 0 && diferencaDias <= 30 && diferencaDias !== 31) {
            var dataFormatada = Utilities.formatDate(dataVencimento, Session.getScriptTimeZone(), "dd/MM/yyyy");
            mensagens.push(
              `Voucher '${aba.getName()}' vence em ${diferencaDias} dias: ${dataFormatada}`
            );
          }
        }
      }
    }
  });

 if (mensagens.length > 0) {
    var conteudoEmail = `
Olá,

    Segue abaixo a lista de vouchers próximos ao vencimento:

${mensagens.join("\n")}

  At.te,
  Yanca Albuquerque!
    `;
    
    Logger.log("📩 E-mail será enviado para: " + emailsDestino);
    Logger.log("📄 Conteúdo do e-mail: \n" + conteudoEmail);
    
    GmailApp.sendEmail(emailsDestino, "Aviso: Vouchers próximos ao vencimento", conteudoEmail);
  }
}