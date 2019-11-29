//function sendEmail(e){
//  var itemResponses = e.response.getItemResponses();
//  var userName = itemResponses[1].getResponse();
//  
//  var html = HtmlService.createTemplateFromFile("email.html");
//  var htmlText = html.evaluate().getContent();
//  var emailTo = e.response.getRespondentEmail();
//  var subject = "Здравствуйте, " + userName + "! Мы получили Вашу анкету!";
//  var textBody = "The Email requires HTML support!";
//  var options = { htmlBody: htmlText };
//  if(emailTo !== undefined){
//   GmailApp.sendEmail(emailTo, subject, textBody, options)};
//}

function onSubmit(e) {
  var itemResponses = e.response.getItemResponses();
  var email = e.response.getRespondentEmail();
  var timestamp = e.response.getTimestamp();
  
  var userLastName = itemResponses[0].getResponse();
  var userFirstName = itemResponses[1].getResponse();
  var userSecondName = itemResponses[2].getResponse();
  var coAuthors = itemResponses[3].getResponse();
  var articleName = itemResponses[4].getResponse();
  var articleFile = itemResponses[5].getResponse();
  var phoneNumber = itemResponses[6].getResponse();
  
  // ----------- Connecting to the spreadsheet ------------ //
  var ss = SpreadsheetApp.openById("1MIcNnm5Z-hHM1C1jLfMHOctqqzLdtBOHQKfX5NDImaQ");
  var sheetResults = ss.getSheetByName("sheet1");
  var lastRow = sheetResults.getLastRow(); // getting the last row
  
  // ----------- Creating 2d arrays of data rows for loops ---------- //
  var lastNames = sheetResults.getRange(2, 1, lastRow, 1).getValues(); // a list of last names
  var firstNames = sheetResults.getRange(2, 2, lastRow, 2).getValues(); // a list of first names
  var secondNames = sheetResults.getRange(2, 3, lastRow, 3).getValues(); // a list of second names
  var emails = sheetResults.getRange(2, 6, lastRow, 6).getValues(); // a list of emails names
  // ----------------------------------------------------------------- //
  var flag = false;   // for indentifying matching rows (last names, first names, last names and emails)
  var indexValue = 0; // for science purposes of course
  
  // ------------------- searching for matches ------------------- //
  for(var i = 0; i < lastNames.length; i++) { // Checking matches in rows
    for(var j = 0; j < lastNames[i].length; j++) {
      if(lastNames[i][j] === userLastName && firstNames[i][j] === userFirstName && secondNames[i][j] === userSecondName && emails[i][j] === email) { // checking matches with last names, first names, optional second names and emails 
        flag = true;
        indexValue = i; }
    }
  }
  
  if(flag) {   // Updating specific data if flag is true
    sheetResults.getRange(indexValue+2, 16).setValue(timestamp);
    if(phoneNumber) {sheetResults.getRange(indexValue+2, 7).setValue(phoneNumber);} // if phone number is true
    sheetResults.getRange(indexValue+2, 12).setValue(1);
    if(coAuthors) {sheetResults.getRange(indexValue+2, 14).setValue(coAuthors);} // if co-Authors is true
    sheetResults.getRange(indexValue+2, 13).setValue(articleName);
    GmailApp.sendEmail(email, 
                       'Здравствуйте, ' + userFirstName + '! Мы получили Вашу анкету!',
                       'The Email requires HTML support',
                       {htmlBody: '<img src="http://95.216.215.138/guap-full-line.png" alt="guap" style="display: block; margin: auto; width: 95%;"><h2>Спасибо за Вашу анкету, ' + userFirstName + '! </h2><p>Ваша анкета для публикации статьи отправлена(обновлена)!</p>'}
                      );
  } else {
    GmailApp.sendEmail(email, 
                       'Здравствуйте, ' + userFirstName + '! Возникли проблемы с анкетой!',
                       'The Email requires HTML support',
                       {htmlBody: '<img src="http://95.216.215.138/guap-full-line.png" alt="guap" style="display: block; margin: auto; width: 95%;"><h2>Доброго времени суток, ' + userFirstName + '! </h2><p>С Вашей анкетой возникли проблемы. Убедитесь в том, что вы уже заполнили форму <a href="https://docs.google.com/forms/d/e/1FAIpQLSfWyhqiKH1YkSg5kjyq4xkCDlN6QSwsInSWDT3ZllbEREsJaQ/viewform">"Заявка на участие"</a>! И проверьте корректность введенных данных. Либо свяжитесь с администратором по электронной почте m.polyak@guap.ru</p>'}
                      );
  }
}