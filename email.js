function sendEmail() {

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dados").activate();

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //pegando o último índice da linha
  var lr = ss.getLastRow();
  
  var templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1,1).getValue();
  
  var cotaEmail = MailApp.getRemainingDailyQuota();
  Logger.log(cotaEmail)
  
  if(lr>cotaEmail){
    Browser.msgBox("Cota de email diário ultrapassada "+(cotaEmail)+" tente novamente amanhã.");
  }
  
  for(var i=2; i<=lr; i++){
    
    var currentEmail = ss.getRange(i, 5).getValue();
    var nome = ss.getRange(i, 3).getValue();
    var assunto = "Setor de Estágio. Confirmação de orientação";
    var messageBody  = templateText.replace("{name}", nome);
    //Logger.log(currentEmail)
    
//    MailApp.sendEmail(currentEmail, assunto, messageBody)
  }
}
