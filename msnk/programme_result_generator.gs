function myProgramResultGenerator() {
  //---Ссылки---
  var table = "1zkLdYvq8qjKB4pOM6h8OwDPHahfrH9n8_GOUtySW-p0" //адрес таблицы
  var folder = '1glnthyQxrflTlEQ0dUtjs2YlQqWUVUA0' // адрес папки для сохранения файла
  
  //---Шапка документа---
  var head = 'Отчет о проведении 72-й МСНК ГУАП'
  var headNext = 'Секция 43. Компьютерных технологий и программной инженерии'
  
  //---Под подпись - Научный руководитель секции---
  var bootom = 'Охтилев М.Ю.'
  
  //---Шапки таблиц---
  var dateTime = ['17 апреля 2019 г., 16-00', 
                  '22 апреля 2019 г., 10-00']
  
  var adress = ['ул. Б. Морская, д. 67, ауд. 52-09', 
                'ул. Б. Морская, д. 67, ауд. 23-10']
  
  //Научный руководитель секции
  var directer = ['д-р техн. наук, проф. М.Ю. Охтилев', 
                  'д-р техн. наук, проф. М.Ю. Охтилев']
  
  //Секретарь
  var secretary = ['ст. преп. М.Д. Поляк', 
                   'д-р техн. наук, проф. С.И. Колесникова']
  
  //---Решения---
  var reportGood = 'Опубликовать доклад в сборнике СНК'
  var reportVeryGood = 'Опубликовать доклад в сборнике СНК; рекомендовать к участию в финале конкурса на лучшую студенческую научную работу ГУАП'
  var reportBad = 'Доклад плохо подготовлен'
  
  
  //---Подключаемся к таблице и создаем файл в нужной папке---
  var docName = 'Список представляемых к публикации докладов';
  var sheet = SpreadsheetApp.openById(table).getActiveSheet();
  var data = sheet.getDataRange().getValues(); // данные этой таблицы 
  var doc = DocumentApp.create(docName); // создание нового документа 
    
  var docFile = DriveApp.getFileById( doc.getId() );
 
  DriveApp.getFolderById(folder).addFile( docFile );
  DriveApp.getRootFolder().removeFile(docFile);  
  
  var body = doc.getBody(); // получение тела документа
  
  //---Стили оформления---
  var styleBody = {};
  styleBody[DocumentApp.Attribute.FONT_FAMILY] = 'Times New Roman'; 
  styleBody[DocumentApp.Attribute.FONT_SIZE] = 10; 
  body.setAttributes(styleBody);

  var styleHead1 = {}; 
  styleHead1[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  styleHead1[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;
  styleHead1[DocumentApp.Attribute.BOLD] = true; 
  
  var styleHead1_1 = {}; 
  styleHead1_1[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  styleHead1_1[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;
  styleHead1_1[DocumentApp.Attribute.BOLD] = false; 
   
  var styleHead1_2 = {};
  styleHead1_2[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  styleHead1_2[DocumentApp.Attribute.BOLD] = true; 
  
  var styleHead2 = {};
  styleHead2[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  styleHead2[DocumentApp.Attribute.BOLD] = false; 
  //----------------------------------------------------------------------------------------------------
  
  var titles = []; // доступ к столбцам по названию 
  for (var i = 0; i < data[0].length; ++i) {
      titles[data[0][i]] = i; // получаю номера столбцов по их названиям 
  }
  
  //---------------------------------------------------
  //Получаем даты выступлений
  var a = []
  for (var i = 1; i < data.length; ++i) {
    if (data[i][titles['Выступление']] != '') {
      a.push(Date.parse(data[i][titles['Выступление']]))
    }
  }
  
  var unique = a.filter( onlyUnique );
  
  unique.sort() //Сортируем по дате от раннего в позднему
  
  //----------------------------------------------
  
  body.appendParagraph(head).setAttributes(styleHead1);
  body.appendParagraph(headNext).setAttributes(styleHead1); //в примере нет выравнивания по центу, но так вроде получше
  body.appendParagraph(' ')
  
  for (var i = 0; i < unique.length; ++i) {
    
    body.appendParagraph('Заседание ' + (i+1)).setAttributes(styleHead1_2)
    body.appendParagraph(dateTime[i] + '				' + adress[i]).setAttributes(styleHead2)
    
    if (directer.length > 1)
      body.appendParagraph('Научный руководитель секции – ' + directer[i]).setAttributes(styleHead2)
    else
      body.appendParagraph('Научный руководитель секции – ' + directer[1]).setAttributes(styleHead2)
    
    body.appendParagraph('Секретарь – ' + secretary[i]).setAttributes(styleHead2)
    body.appendParagraph(' ')
    body.appendParagraph('Список докладов').setAttributes(styleHead2)
    body.appendParagraph(' ')
    
    var table = body.appendTable();   //Создаем новую таблицу
    var tr1 = table.appendTableRow(); //Создаем шапку для нее
    tr1.appendTableCell('№ п/п').getChild(0).asParagraph().setAttributes(styleHead1);
    tr1.appendTableCell('Фамилия и инициалы докладчика, название доклада').getChild(0).asParagraph().setAttributes(styleHead1);
    tr1.appendTableCell('Статус (магистр / студент)').getChild(0).asParagraph().setAttributes(styleHead1);
    tr1.appendTableCell('Решение').getChild(0).asParagraph().setAttributes(styleHead1);

    var serialNumber = 1 //Номер по порядку
    for (var j = 1; j < data.length; ++j) {
      if (Date.parse(data[j][titles['Выступление']]) === unique[i]) {
        var tr = table.appendTableRow()
        
        tr.appendTableCell(serialNumber++).getChild(0).asParagraph().setAttributes(styleHead1_1);
        
        tr.appendTableCell(data[j][titles['Фамилия']] + ' ' + 
            data[j][titles['Имя']] + ' ' +
            data[j][titles['Отчество']] + '. ' +
            data[j][titles['Тема']]).getChild(0).asParagraph().setAttributes(styleHead1_1);
        
        var tempGroup = data[j][titles['Группа']]
        if (tempGroup[tempGroup.length-1] === 'M' || tempGroup[tempGroup.length-1] === 'М') {
          tr.appendTableCell('Магистрант гр.' + data[j][titles['Группа']]).getChild(0).asParagraph().setAttributes(styleHead1_1);
        }
        else {
          tr.appendTableCell('Студент гр.' + tempGroup).getChild(0).asParagraph().setAttributes(styleHead1_1);
        }

        if (data[j][titles['Рекомендация']] === 1 || data[j][titles['Рекомендация']] === 'Да') {
          tr.appendTableCell(reportGood).getChild(0).asParagraph().setAttributes(styleHead1_1);
        }
        else if (data[j][titles['Рекомендация']] === 2) {
          tr.appendTableCell(reportVeryGood).getChild(0).asParagraph().setAttributes(styleHead1_1);
        }
        else {
          tr.appendTableCell(reportBad).getChild(0).asParagraph().setAttributes(styleHead1_1);
        }
      }
     
    }
    table.setColumnWidth(0, 30)
    table.setColumnWidth(1, 250)
    table.setColumnWidth(2, 70)
    table.setColumnWidth(3, 100)
  }
  body.appendParagraph(' ')
  body.appendParagraph('Научный руководитель секции                                    _________________ / ' + bootom).setAttributes(styleHead2);
}

//Получаем уникальные значения
function onlyUnique(value, index, self) { 
    return self.indexOf(value) === index; 
}
