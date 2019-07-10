function generate_conference() { 
	//Ссылки 
	var table = "1zkLdYvq8qjKB4pOM6h8OwDPHahfrH9n8_GOUtySW-p0" //адрес таблицы  
	var folder = '1tQk-2NEQfBAl31LqrimTi_qW2ubGeJez'   //Папка для сохранения

	//Шапка документа 
	var head = 'Программа 72-й МСНК ГУАП' 
	var headNext = 'по кафедре № 43 компьютерных технологий и программной инженерии' 

	//Подключаемся к таблице и создаем файл в нужной папке 
	var docName = 'Программа конференции'; 
	var sheet = SpreadsheetApp.openById(table).getActiveSheet(); 
	var data = sheet.getDataRange().getValues(); // данные этой таблицы 
	var doc = DocumentApp.create(docName); // создание нового документа 

	var docFile = DriveApp.getFileById( doc.getId() ); 

	DriveApp.getFolderById(folder).addFile( docFile ); 
	DriveApp.getRootFolder().removeFile(docFile); 

	var body = doc.getBody(); // получение тела документа 

	//Стили оформления 
	var styleList = {}; 
	styleList[DocumentApp.Attribute.FONT_SIZE] = 14; 
	styleList[DocumentApp.Attribute.INDENT_START] = 36; 
	styleList[DocumentApp.Attribute.INDENT_FIRST_LINE] = 18; 
	styleList[DocumentApp.Attribute.SPACING_AFTER] = 10; 

	var styleHead = {}; 
	styleHead[DocumentApp.Attribute.FONT_FAMILY] = 'Times New Roman'; 
	styleHead[DocumentApp.Attribute.FONT_SIZE] = 14; 
	styleHead[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER; 
	styleHead[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER; 
	styleHead[DocumentApp.Attribute.BOLD] = true; 
	styleHead[DocumentApp.Attribute.ITALIC] = true; 

	var styleBody1 = {}; 
	styleBody1[DocumentApp.Attribute.FONT_SIZE] = 12; 
	styleBody1[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT; 
	styleBody1[DocumentApp.Attribute.INDENT_START] = 36; 
	styleBody1[DocumentApp.Attribute.INDENT_FIRST_LINE] = 18; 
	styleBody1[DocumentApp.Attribute.BOLD] = true; 
	styleBody1[DocumentApp.Attribute.ITALIC] = false; 

	var styleBody2 = {}; 
	styleBody2[DocumentApp.Attribute.FONT_SIZE] = 12; 
	styleBody2[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT; 
	styleBody2[DocumentApp.Attribute.INDENT_START] = 18; 
	styleBody2[DocumentApp.Attribute.INDENT_FIRST_LINE] = 64; 
	styleBody2[DocumentApp.Attribute.BOLD] = false; 

	var titles = []; // доступ к столбцам по названию 
	for (var i = 0; i < data[0].length; ++i) { 
		titles[data[0][i]] = i; // получаю номера столбцов по их названиям 
	} 

	body.appendParagraph(head).setAttributes(styleHead); 
	body.appendParagraph(headNext).setAttributes(styleHead); 
	body.appendParagraph(' ') 
	body.appendParagraph('Заседание 1: 17 апреля 2019 (среда), 16:00, ауд. 23-10 БМ').setAttributes(styleBody1) 
	body.appendParagraph('Председатель – к.т.н., доц. Ключарев Александр Анатольевич').setAttributes(styleBody2) 
	body.appendParagraph('Секретарь – ст. преп. Поляк М.Д.').setAttributes(styleBody2) 
	body.appendParagraph(' ') 

	// отбрасывание строки с названиями столбцов 
	data.splice(0,1); 

	// преобразование номеров групп к общему виду 
	for (var i = 0; i < data.length; ++i) { 
		data[i][titles['Группа']] = data[i][titles['Группа']].toString().replace(/m|M|м/g,'М') 
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
	sortByGroup(data, titles['Группа']); 

	// добавление студентов с их группами в документ списком 
	for (var i = 0; i < data.length; ++i) { 
		if (data[i][titles['Группа']][data[i][titles['Группа']].length-1] != 'М') 
			body.appendListItem( 
			data[i][titles['Фамилия']] + ' ' + 
			data[i][titles['Имя']] + ' ' + 
			data[i][titles['Отчество']] + ', группа ' + 
			data[i][titles['Группа']] + '\r' + 
			data[i][titles['Тема']]).setAttributes(styleList); 
	} 

	// добавление магистрантов с их группами в документ списком 
	for (var i = 0; i < data.length; ++i) { 
		if (data[i][titles['Группа']][data[i][titles['Группа']].length-1] === 'М') 
			body.appendListItem( 
			data[i][titles['Фамилия']] + ' ' + 
			data[i][titles['Имя']] + ' ' + 
			data[i][titles['Отчество']] + ', группа ' + 
			data[i][titles['Группа']] + '\r' + 
			data[i][titles['Тема']]).setAttributes(styleList); 
	} 
}