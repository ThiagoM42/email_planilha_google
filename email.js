
var formID = '1E214puIOqI8aTfWg1UkywwO4eXbdp_5xw6gNaVJROaE';
var ssID = '1VcHtSwKJetp6JBD-Rw6FyZQ-NFHOwbIwAllpt5MlX18';

var wsData = SpreadsheetApp.openById(ssID).getSheetByName("dados_professores");
var form = FormApp.openById(formID);
var item = form.getItemById(1409839857);

function sendEmail() {

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dados").activate();

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //pegando o último índice da linha
  var lr = ss.getLastRow();
  
  var templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1,1).getValue();   
  
  

    for(var i=2; i<=lr; i++){
      
      var currentEmail = ss.getRange(i, 10).getValue();
      Logger.log(currentEmail);
      var nome = ss.getRange(i, 3).getValue();
      var nomeProfessor = ss.getRange(i, 9).getValue();
      var curso = ss.getRange(i, 7).getValue();
      var assunto = "Setor de Estágio. Confirmação de orientação";
      var messageBody  = templateText.replace("{name}", nome).replace("{nome_professor}", nomeProfessor).replace("{curso}", curso);
      
      var confirmado =ss.getRange(i, 15).getValue();
      var email_ja_enviado = ss.getRange(i, 14).getValue();
      
      if(!confirmado && !email_ja_enviado && currentEmail){
        MailApp.sendEmail(currentEmail, assunto, messageBody)
      }
    }//end for
     Browser.msgBox("Email enviado de confirmação enviado para todos os professores");
     emailConfirm();
  
}

function emailConfirm(){
  
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dados").activate();
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var lr = ss.getLastRow();
  
  for(var i=2; i<=lr; i++){
    var currentEmail = ss.getRange(i, 10).getValue();
    
    var qtd_email = ss.getRange(i, 14).getValue();
    qtd_email = qtd_email + 1;
    
    var confirmado =ss.getRange(i, 15).getValue();
      var email_ja_enviado = ss.getRange(i, 14).getValue();
      
      if(!confirmado && !email_ja_enviado && currentEmail){
      ss.getRange(i, 14).setValue(qtd_email).setBackground('green').setFontColor('white').setFontFamily('bold');
    }
  }
  
}


function addLista() {
  
  var values = wsData.getRange(1, 2, wsData.getLastRow(), 1).getValues();
  values = values.filter(function(val){return val != ""})
  item.asListItem().setChoiceValues(values);
    
}


