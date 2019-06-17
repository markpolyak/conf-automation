function generateGoogleDocs() {
    var docName = 'Темы докладов';
    var sheet = SpreadsheetApp.openById("1RaxHX9p1I8qCS2JwQajZ0E7y3_wtltYrJak0dioeoks").getActiveSheet(); // текущая таблица
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

    // отбрасывание строки с названиями столбцов 
    data.splice(0,1);

    // преобразование номеров групп к общему виду 
    for (var i = 0, len = data.length; i < len; ++i) {
        data[i][titles['Номер группы']] = data[i][titles['Номер группы']].toString() .replace(/m|M|м/g,'М')
        .replace(/z|з|З/g,'Z')
        .replace(/k|K|к/g,'К')
        .replace(/v|V|в/g,'В'); 
    }

    // сортировка студентов по номеру группы 
    function sortByGroup(arr, index) {
        function sortFunc(val1, val2) {
            val1 = val1[index];
            val2 = val2[index];
            return (val1 === val2) ? 0 : (val1 < val2) ? -1 : 1
        }
        data.sort(sortFunc); 
    }
    //sortByGroup(data, titles['Номер группы']);

    // добавление студентов с их группами в документ списком 
    for (var i = 0, len = data.length; i < len; ++i) {
        var listItem = body.appendListItem(data[i][titles['Фамилия']] + ' ' + 
            data[i][titles['Имя']] + ' ' +
            data[i][titles['Отчество']] + ', группа ' + 
            data[i][titles['Номер группы']]).setAttributes(styleList);
    }
    // добавление параграфов с названиями тем
    for (var i = 0, j = 1, len = data.length; i < len; ++i, ++j) {
        body.insertParagraph(i+j+1, data[i][titles['Тема доклада']]).setAttributes(stylePar); 
    }

    var header1_1 = "Темы докладов для программы 72-й МСНК ГУАП ";
    var header1_2 = "по кафедре No 43 компьютерных технологий и программной инженерии"; var empty = "";
    var header2 = "Секция каф.43. «компьютерных технологий и программной инженерии»"; var header3_1 = "Научный руководитель секции – Охтилев Михаил Юрьевич"
    var header3_2 = "зав. кафедрой, д-р техн. наук, профессор"
    var header3_3 = "Зам. научного руководителя секции – Ключарев А.А."
    var header3_4 = "канд. техн. наук, доцент";
    // вставка шапки документа
    body.insertParagraph(0, header3_4).setAttributes(styleHead3); 
    body.insertParagraph(0, header3_3).setAttributes(styleHead3); 
    body.insertParagraph(0, header3_2).setAttributes(styleHead3); 
    body.insertParagraph(0, header3_1).setAttributes(styleHead3);
    body.insertParagraph(0, header2).setAttributes(styleHead2); 
    body.insertParagraph(0, empty).setAttributes(styleHead1); 
    body.insertParagraph(0, header1_2).setAttributes(styleHead1); 
    body.insertParagraph(0, header1_1).setAttributes(styleHead1);
}
