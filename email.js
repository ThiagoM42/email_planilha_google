function sendEmail() {

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dados").activate();

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //pegando o último índice da linha
  var lr = ss.getLastRow();
  
  var templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1,1).getValue();
  
  var cotaEmail = MailApp.getRemainingDailyQuota();
  Logger.log(cotaEmail);
  
  if(lr>cotaEmail){
    Browser.msgBox("Cota de email diário ultrapassada "+(cotaEmail)+" tente novamente amanhã.");
  }
  else{ 
    for(var i=2; i<=lr; i++){
      
      var currentEmail = ss.getRange(i, 10).getValue();
      var nome = ss.getRange(i, 3).getValue();
      var nomeProfessor = ss.getRange(i, 9).getValue();
      var curso = ss.getRange(i, 7).getValue();
      var assunto = "Setor de Estágio. Confirmação de orientação";
      var messageBody  = templateText.replace("{name}", nome).replace("{nome_professor}", nomeProfessor).replace("{curso}", curso);
      
      var confirmado =ss.getRange(i, 15).getValue();
      
      if(!confirmado){
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
    var qtd_email = ss.getRange(i, 14).getValue();
    qtd_email = qtd_email + 1;
    
    var confirmado =ss.getRange(i, 15).getValue();
      
    if(!confirmado){
      ss.getRange(i, 14).setValue(qtd_email).setBackground('green').setFontColor('white').setFontFamily('bold');
    }
  }
  
}


