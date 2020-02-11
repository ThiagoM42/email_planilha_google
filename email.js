function sendEmail() {

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dados").activate();

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //pegando o último índice da linha
  var lr = ss.getLastRow();
  
  var templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1,1).getValue();
  
  var cotaEmail = MailApp.getRemainingDailyQuota();
  
  if(lr>cotaEmail){
    Browser.msgBox("Cota de email diário ultrapassada "+(cotaEmail)+" tente novamente amanhã.");
  }
  else{ 
    for(var i=2; i<=lr; i++){
      
      var currentEmail = ss.getRange(i, 5).getValue();
      var nome = ss.getRange(i, 3).getValue();
      var assunto = "Setor de Estágio. Confirmação de orientação";
      var messageBody  = templateText.replace("{name}", nome);
      var emailEnviado =ss.getRange(i, 6).getValue();
      
      if(emailEnviado){
        MailApp.sendEmail(currentEmail, assunto, messageBody)
      }
    }//end for
     Browser.msgBox("Email enviado de confirmação enviado para todos os professores");
     emailConfirm();
  } //end else
}

function emailConfirm(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dados").activate();
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var lr = ss.getLastRow();
  
  for(var i=2; i<=lr; i++){
    ss.getRange(i, 6).setValue("Sim");
  }
  
}


