function generateGoogleDocs() {
    var docName = 'Список представляемых к публикации докладов';
    var sheet = SpreadsheetApp.openById("1zkLdYvq8qjKB4pOM6h8OwDPHahfrH9n8_GOUtySW-p0").getActiveSheet(); // текущая таблица
    // var sheet = SpreadsheetApp.openById("1zkLdYvq8qjKB4pOM6h8OwDPHahfrH9n8_GOUtySW").getSheetByName('Участники'); // текущая таблица
    var data = sheet.getDataRange().getValues(); // данные этой таблицы 
    // data[0].length - количество полей: 10
    // data.length - количество строк: 40
    
    var doc = DocumentApp.create(docName); // создание нового документа 
    // перемещение документа в нужную папку (см. https://stackoverflow.com/questions/31739653/create-a-google-doc-file-directly-in-a-google-drive-folder )
    var docFile = DriveApp.getFileById( doc.getId() );
    DriveApp.getFolderById('1glnthyQxrflTlEQ0dUtjs2YlQqWUVUA0').addFile( docFile );
    DriveApp.getRootFolder().removeFile(docFile);
    //
    var body = doc.getBody(); // получение тела документа
    var styleBody = {};
    styleBody[DocumentApp.Attribute.FONT_FAMILY] = 'Times New Roman'; 
    styleBody[DocumentApp.Attribute.FONT_SIZE] = 14; 
    body.setAttributes(styleBody);
    // стили для списка
    var styleList = {}; styleList[DocumentApp.Attribute.INDENT_START] = 36; 
    styleList[DocumentApp.Attribute.INDENT_FIRST_LINE] = 18; 
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
    var styleHead2 = {};
    styleHead2[DocumentApp.Attribute.FONT_FAMILY] = 'Times New Roman';
    styleHead2[DocumentApp.Attribute.FONT_SIZE] = 12;
    styleHead2[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
    styleHead2[DocumentApp.Attribute.BOLD] = true; 
    styleHead2[DocumentApp.Attribute.ITALIC] = true; 
    styleHead2[DocumentApp.Attribute.SPACING_AFTER] = 6; 
    styleHead2[DocumentApp.Attribute.INDENT_START] = 35,43307086614; 
    styleHead2[DocumentApp.Attribute.INDENT_FIRST_LINE] = 35,43307086614; 
    var styleHead3 = {};
    styleHead3[DocumentApp.Attribute.FONT_FAMILY] = 'Times New Roman'; 
    styleHead3[DocumentApp.Attribute.FONT_SIZE] = 12; 
    styleHead3[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT; 
    styleHead3[DocumentApp.Attribute.INDENT_START] = 56.69291338583; 
    styleHead3[DocumentApp.Attribute.INDENT_FIRST_LINE] = 56.69291338583;

    var titles = []; // доступ к столбцам по названию 
    for (var i = 0, len = data[0].length; i < len; ++i) {
        titles[data[0][i]] = i; // получаю номера столбцов по их названиям 
    }
    Logger.log(data[0].length)
    Logger.log(data[0])
    Logger.log(titles)

    // отбрасывание строки с названиями столбцов 
    data.splice(0,1);

    // добавление студентов с их группами в документ списком 
    for (var i = 0, len = data.length; i < len; ++i) {
        if (data[i][titles['Электронная версия']] == 1) {
            var listItem = body.appendListItem(data[i][titles['Фамилия']] + '\xa0' + 
                data[i][titles['Имя']].substring(0,1) + '.' + '\xa0' +
                data[i][titles['Отчество']].substring(0,1) + '. ' + 
                data[i][titles['Название статьи']]).setAttributes(stylePar);
        }
    }

    var header1_1 = "Список представляемых к публикации докладов";
    var header1_2 = "Кафедра № 43 компьютерных технологий и программной инженерии"; var empty = "";
    var header2_1 = "Поляк Марк Дмитриевич"
    var header2_2 = "e-mail: m.polyak@guap.ru"
    var header2_3 = "тел.: +7-XXX-XXX-XXXX"
    // вставка шапки документа
    body.insertParagraph(0, header2_3).setAttributes(styleHead3); 
    body.insertParagraph(0, header2_2).setAttributes(styleHead3); 
    body.insertParagraph(0, header2_1).setAttributes(styleHead3);
    body.insertParagraph(0, header1_2).setAttributes(styleHead3); 
    body.insertParagraph(0, empty).setAttributes(styleHead1); 
    body.insertParagraph(0, header1_1).setAttributes(styleHead1);
}
