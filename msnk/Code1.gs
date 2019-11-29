function sendEmail(e){
  var itemResponses = e.response.getItemResponses(); // all user responses from form
  var adminEmail = "insidious_kun@yahoo.com";
  var userName = itemResponses[1].getResponse();     // user's first name from form
  var userLastName = itemResponses[0].getResponse(); // user's last name from form
  //-----------------------------------------------------------------------
  var html = HtmlService.createTemplateFromFile("email.html");
  var htmlText = html.evaluate().getContent();
  var emailTo = e.response.getRespondentEmail();     // user's email from form
  var submittedTime = e.response.getTimestamp();     // timestamp of submitted form
  var subject = "Здравствуйте, " + userName + "! Мы получили Вашу заявку!";
  var textBody = "The Email requires HTML support!";
  var options = { htmlBody: htmlText };
  if(emailTo !== undefined){
    GmailApp.sendEmail(emailTo, subject, textBody, options)
    GmailApp.sendEmail(adminEmail, "Студент оставил(обновил) заявку на участие.", "Студент " + userLastName + " " + userName + " добавил(обновил) свою заявку на участие. \nВремя заявки: " + submittedTime)
  };
}

function onSubmitForm(e){
  // --------- User input from form ---------- //
  var itemResponses = e.response.getItemResponses();
  var email = e.response.getRespondentEmail();
  var timestamp = e.response.getTimestamp();
  var userLastName = itemResponses[0].getResponse();
  var userFirstName = itemResponses[1].getResponse();
  var userSecondName = itemResponses[2].getResponse();
  var reportName = itemResponses[3].getResponse();
  var leaderName = itemResponses[4].getResponse();
  var phoneNumber = itemResponses[5].getResponse();
  var organization = itemResponses[6].getResponse();
  var groupNumber;
  if(itemResponses[7] !== undefined) {
    groupNumber = itemResponses[7].getResponse();
  } else { groupNumber = null; }

  // ----------- Connecting to the spreadsheet ------------ //
  var ss = SpreadsheetApp.openById("1MIcNnm5Z-hHM1C1jLfMHOctqqzLdtBOHQKfX5NDImaQ");
  var sheetResults = ss.getSheetByName("sheet1");
  var lastRow = sheetResults.getLastRow(); // getting the last row
  // ----------- Creating 2d arrays of data rows for loops ---------- //
  var lastNames = sheetResults.getRange(2, 1, lastRow, 1).getValues(); // a list of last names
  var firstNames = sheetResults.getRange(2, 2, lastRow, 2).getValues(); // a list of first names
  var secondNames = sheetResults.getRange(2, 3, lastRow, 3).getValues(); // a list of second names
  var groupNumbersAndOrganizations = sheetResults.getRange(2, 4, lastRow, 4).getValues(); // a list of groups and organizations
  // ----------------------------------------------------------------- //
  var flag = false;   // for indentifying matching rows (last names, first names, last names and optionaly group numbers or organizations)
  var indexValue = 0; // for science purposes of course
  
  // ------------------- The logic itself ------------------- //
 
      for(var i = 0; i < lastNames.length; i++) { // Checking matches in rows
        for(var j = 0; j < lastNames[i].length; j++) {
            if(lastNames[i][j] === userLastName && firstNames[i][j] === userFirstName && secondNames[i][j] === userSecondName && organization === 'ГУАП' && groupNumbersAndOrganizations[i][j] == groupNumber) { // if ГУАП можно объединить с другим if с помощью (userSecondName &&...)||(!userSecondName &&...)
              flag = true;
              indexValue = i; }
            else if(lastNames[i][j] === userLastName && firstNames[i][j] === userFirstName && secondNames[i][j] === userSecondName && groupNumbersAndOrganizations[i][j] === organization) { // if other organization можно объединить с другим else if с помощью (userSecondName &&...)||(!userSecondName &&...)
              flag = true;
              indexValue = i; }
        }
      }

    //var form = FormApp.openById('1Me7zh047mUXoCqq8g9GSSiRgvef5Mcu4BzOEczQaZ4U');
    if(flag) {   // Updating specific data if flag is true
      //form.setConfirmationMessage('Thanks for responding!');
      sheetResults.getRange(indexValue+2, 15).setValue(timestamp);
      if(phoneNumber) {sheetResults.getRange(indexValue+2, 7).setValue(phoneNumber);} // if phone number is true
      sheetResults.getRange(indexValue+2, 5).setValue(reportName);
      sheetResults.getRange(indexValue+2, 6).setValue(email);
      sheetResults.getRange(indexValue+2, 8).setValue(leaderName);
    } else {     // if no matches found inserting every value
      //form.setConfirmationMessage('Unfortunately, we couldnt find any user information in our database!');
      sheetResults.getRange(lastRow+1, 15).setValue(timestamp);
      sheetResults.getRange(lastRow+1, 6).setValue(email);
      sheetResults.getRange(lastRow+1, 1).setValue(userLastName);
      sheetResults.getRange(lastRow+1, 2).setValue(userFirstName);
      sheetResults.getRange(lastRow+1, 3).setValue(userSecondName);
      sheetResults.getRange(lastRow+1, 5).setValue(reportName);
      sheetResults.getRange(lastRow+1, 8).setValue(leaderName);
      if(organization === 'ГУАП') {sheetResults.getRange(lastRow+1, 4).setValue(groupNumber);} else {
                                   sheetResults.getRange(lastRow+1, 4).setValue(organization);}
      sheetResults.getRange(lastRow+1, 7).setValue(phoneNumber);
    }
  }