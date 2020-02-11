function sendEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //pegando o último índice da linha
  var lr = ss.getLastRow();
  
  for(var i=2; i<=lr; i++){
    
    var currentEmail = ss.getRange(i, 5).getValue();
    var nome = ss.getRange(i, 3).getValue();
    //Logger.log(currentEmail)

    var messageBody = '';

    var templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1,1).getValue();
    
    //MailApp.sendEmail(currentEmail, "Email de confirmação do FORM: ", "Olá "+nome)
  }
}
