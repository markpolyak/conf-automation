//Строку с телефоном заполнить  
function generateGoogleDocs() {
    var docName = 'Список представляемых к публикации докладов';
    var sheet = SpreadsheetApp.openById("1zkLdYvq8qjKB4pOM6h8OwDPHahfrH9n8_GOUtySW-p0").getActiveSheet(); // текущая таблица
    var data = sheet.getDataRange().getValues(); // данные этой таблицы 
    
    var doc = DocumentApp.create(docName); // создание нового документа 
    var docFile = DriveApp.getFileById( doc.getId() );
    DriveApp.getFolderById('1jYIKuW19jtazhL6j4Tw8OjR4ySlD7gHK').addFile( docFile );

    var body = doc.getBody(); // получение тела документа
    var styleBody = {};
    styleBody[DocumentApp.Attribute.FONT_FAMILY] = 'Times New Roman'; 
    styleBody[DocumentApp.Attribute.FONT_SIZE] = 14; 
    body.setAttributes(styleBody);
    // стили для названий тем
    var stylePar = {}; 
    stylePar[DocumentApp.Attribute.INDENT_FIRST_LINE] = 36; 
    stylePar[DocumentApp.Attribute.INDENT_START] = 0;
    stylePar[DocumentApp.Attribute.SPACING_AFTER] = 10;
    stylePar[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.JUSTIFY;
    // стли для шапки
    var styleHead1 = {}; 
    styleHead1[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    styleHead1[DocumentApp.Attribute.BOLD] = true; 
    styleHead1[DocumentApp.Attribute.ITALIC] = true;
    var styleHead3 = {};
    styleHead3[DocumentApp.Attribute.FONT_SIZE] = 12; 
    styleHead3[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT; 
    styleHead3[DocumentApp.Attribute.INDENT_START] = 56.90551181102; 
    styleHead3[DocumentApp.Attribute.INDENT_FIRST_LINE] = 56.90551181102;

    var titles = []; // доступ к столбцам по названию 
    for (var i = 0, len = data[0].length; i < len; ++i) {
        titles[data[0][i]] = i; // получаю номера столбцов по их названиям 
    }

    // отбрасывание строки с названиями столбцов 
    data.splice(0,1);

    // добавление студентов с их группами в документ списком 
    for (var i = 0, len = data.length; i < len; ++i) {
        if (data[i][titles['Электронная версия']] == 1 || data[i][titles['Электронная версия']] == 'Да') {
          if (data[i][titles['Распечатка']] == 1 || data[i][titles['Распечатка']] == 'Да')
          {
            var listItem = body.appendListItem(data[i][titles['Фамилия']] + '\xa0' + 
                data[i][titles['Имя']].substring(0,1) + '.' + '\xa0' +
                data[i][titles['Отчество']].substring(0,1) + '. ' + 
                data[i][titles['Название статьи']]).setAttributes(stylePar);
          }
        }
    }

    var header1_1 = "Список представляемых к публикации докладов";
    var header1_2 = "Кафедра № 43 компьютерных технологий и программной инженерии";
    var empty = "";
    var header2_1 = "Поляк Марк Дмитриевич"
    var header2_2 = "e-mail: m.polyak@guap.ru"
    var header2_3 = "тел.: +7-999-XXX-XXXX"
    // вставка шапки документа
    body.insertParagraph(0, header2_3).setAttributes(styleHead3); 
    body.insertParagraph(0, header2_2).setAttributes(styleHead3); 
    body.insertParagraph(0, header2_1).setAttributes(styleHead3);
    body.insertParagraph(0, header1_2).setAttributes(styleHead3); 
    body.insertParagraph(0, empty).setAttributes(styleHead1); 
    body.insertParagraph(0, header1_1).setAttributes(styleHead1);
}

//Функция обработки события 
function myFunction() {
  
  var form = FormApp.openById("1SYwEX6QAkujXK2dFwSBoz5WTW6ZOFT3BdwwT9GvF5lE");
  
  var formResponses = form.getResponses();
for (var i = 0; i < formResponses.length; i++) {
  var formResponse = formResponses[i];
  var itemResponses = formResponse.getItemResponses();
}
  var resp = itemResponses[itemResponses.length-1].getResponse();
  for(var j = 0; j < 3; j++)
  {
    if(resp[j] == "Генератор программы конференции")
    {
      Logger.log(resp[j]);
    }
    if(resp[j] == "Генератор отчета")
    {
      Logger.log(resp[j]);
    }
    if(resp[j] == "Генератор списка представляемых к публикации статей")
    {
      Logger.log(resp[j]);
      generateGoogleDocs();
    }
  }
}
