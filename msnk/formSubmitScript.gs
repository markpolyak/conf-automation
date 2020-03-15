function onSubmitForm(e) {
  var sheet = SpreadsheetApp.openById("1vVtLekc9ezWPGq1eRvlH18Fe-7645aVStu7kBmU_jjc").getSheetByName("test");
  var form = FormApp.getActiveForm();
  check(sheet,e);
  form.setConfirmationMessage("Спасибо");
}
function check(sheet,e){
  var itemResponses = e.response.getItemResponses();
  
  //Дата и время заполнения формы
  var date = new Date();  
  var dateResponse = Utilities.formatDate (date, 'GMT+3', 'dd.MM.yyyy HH:mm:ss');
  
  //Почта отправителя формы
  var mail = e.response.getRespondentEmail();
  
  //Проверка на совпадение и обновление данных
  for(var i = 2; i <= sheet.getLastRow(); i++){
    var ch = sheet.getRange(i,1).getValue();
    if(mail == ch){
      var ch = sheet.getRange(i,2).getValue();
      var surname = itemResponses[0].getResponse();
      if(surname == ch){
        var ch = sheet.getRange(i,3).getValue();
        var name = itemResponses[1].getResponse();
        if(name == ch){
          var ch = sheet.getRange(i,4).getValue();
          var patr = itemResponses[2].getResponse();
          if(patr == ch){
            sheet.getRange(i,14).setValue(itemResponses[3].getResponse());//coau
            sheet.getRange(i,13).setValue(itemResponses[4].getResponse());//article name
            sheet.getRange(i,12).setValue(1);
            sheet.getRange(i,7).setValue(itemResponses[6].getResponse());
            sheet.getRange(i,15).setValue(dateResponse);
          }
        }
      }
    }
  }
}