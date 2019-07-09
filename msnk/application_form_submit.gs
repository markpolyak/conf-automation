function onSubmitForm(e){
  Logger.log("onSubmitForm running");
  
  var ss = SpreadsheetApp.openById("1cnsaad-s2Bvz4LZjObw04-hz2Gi7NiaynRxaBGGBR1I"); // Replace with your Spreadsheet Id
  var sheet = ss.getSheetByName("Sheet1");                                          // Replace with your Sheet name
  var row = sheet.getLastRow();
  var itemResponses = e.response.getItemResponses();
  var emailAddress = e.response.getRespondentEmail();
  
  //Check
  var endFun = 0;
  endFun = check(itemResponses, emailAddress, sheet, row); 
  
  if (endFun == 0) {Logger.log("Function error")}
  else if (endFun == 1) {Logger.log("Response has been updated")}
  else if (endFun == 2) {Logger.log("Response added")};
  
}

//Add the response
function addInTable(itemResponses, emailAddress, sheet, row) {
  var date = new Date();  
  var dateString = Utilities.formatDate (date, 'GMT+3', 'dd.MM.yyyy HH:mm:ss');
  
  for (var i = 0; i < itemResponses.length; i++) {
    var response = itemResponses[i].getResponse();
    sheet.getRange(row, i+2).setValue(response);
  }
  //Add the date of the response
  sheet.getRange(row, 1).setValue(emailAddress);
  sheet.getRange(row, i+2).setValue(dateString);
}


//Sequential check for the coincidence of surname, name, patronymic and Institute
function check(itemResponses, emailAddress, sheet, row) {
  Logger.log("Check run");
  
  var surnameResponse = itemResponses[1].getResponse(); 
  Logger.log("surnameResponse = " + surnameResponse);
  for (var i = 2; i <= row; i++) {
    var check = sheet.getRange(i, 2).getValue();
    if (surnameResponse == check) {
      var nameResponse = itemResponses[2].getResponse(); 
      Logger.log("nameResponse = " + nameResponse);
      var check = sheet.getRange(i, 3).getValue();
      if (nameResponse == check) {
        var patronymicResponse = itemResponses[3].getResponse();
        Logger.log("patronymicResponse = " + patronymicResponse);
        var check = sheet.getRange(i, 4).getValue();
        if (patronymicResponse == check) {
          var instituteResponse = itemResponses[6].getResponse();
          Logger.log("instituteResponse = " + instituteResponse);
          if (instituteResponse == "ГУАП") {
            var groupResponse = itemResponses[7].getResponse();
            Logger.log("groupResponse = " + groupResponse);
            var check = sheet.getRange(i, 8).getValue();
            if (groupResponse == check) {
              addInTable(itemResponses, emailAddress, sheet, i);
              Logger.log("Check end");
              return 1;
            }
          } else {
            var check = sheet.getRange(i, 7).getValue();
            if(instituteResponse == check){
              addInTable(itemResponses, emailAddress, sheet, i);
              Logger.log("Check end");
              return 1;
            }
          }
        }
      }
    }
  }
 
  addInTable(itemResponses, emailAddress, sheet, row+1); 
  Logger.log("Check end");
  return 2;
}
