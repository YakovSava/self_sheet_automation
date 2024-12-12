/**
 * Функция для группировки данных в колонках B и A,
 * а также обновления значений в колонке M на основе совпадений.
 */
function groupData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('current'); // Убедись, что название листа корректно

  // === Группировка имён работников в колонке B:B и вывод в Q начиная с Q5 ===
  const employeeRange = sheet.getRange('B2:B'); // Начинаем с B2, исключая заголовок
  const employeeValues = employeeRange.getValues().flat().filter(name => name && name.toString().trim() !== '');
  const uniqueEmployees = [...new Set(employeeValues)].sort();

  // Очищаем предыдущие данные в колонке Q начиная с Q5
  sheet.getRange('Q5:Q').clearContent();

  // Записываем уникальные имена работников в Q5 и ниже
  if (uniqueEmployees.length > 0) {
    sheet.getRange(5, 17, uniqueEmployees.length, 1).setValues(uniqueEmployees.map(name => [name]));
  }

  // === Группировка названий компаний в колонке A:A и вывод в U начиная с U2 ===
  const companyRange = sheet.getRange('A2:A'); // Начинаем с A2, исключая заголовок
  const companyValues = companyRange.getValues().flat().filter(name => name && name.toString().trim() !== '');
  const uniqueCompanies = [...new Set(companyValues)].sort();

  // Очищаем предыдущие данные в колонке U начиная с U2
  sheet.getRange('U2:U').clearContent();

  // Записываем уникальные названия компаний в U2 и ниже
  if (uniqueCompanies.length > 0) {
    sheet.getRange(2, 21, uniqueCompanies.length, 1).setValues(uniqueCompanies.map(name => [name]));
  }

  // === Обновление значений в колонке M на основе совпадений с колонкой U ===
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // Если нет данных, выходим

  // Получаем данные из колонок A и X, начиная с 2 строки
  const companyData = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat(); // A2:A
  const xData = sheet.getRange(2, 24, lastRow - 1, 1).getValues().flat(); // X2:X

  // Получаем объект сопоставления компаний и их значений из X
  const companyToXMap = {};
  uniqueCompanies.forEach((company, index) => {
    // Предполагаем, что значения в X соответствуют строкам U (начиная с U2)
    // То есть, X2 соответствует U2, X3 - U3 и т.д.
    const xValue = sheet.getRange(2 + index, 24).getValue(); // X2, X3, ...
    companyToXMap[company] = xValue;
  });

  // Создаем массив для записи в колонку M
  const mValues = companyData.map(company => {
    if (companyToXMap.hasOwnProperty(company)) {
      return [companyToXMap[company]];
    } else {
      return [''];
    }
  });

  // Записываем значения в колонку M начиная с M2
  sheet.getRange(2, 13, mValues.length, 1).setValues(mValues);
}

/**
 * Функция, которая срабатывает при изменении ячеек в таблице.
 * Проверяет, изменена ли ячейка H2 на 'Убрать в архив (автоматизация)'.
 * Если да, запускает процесс архивации.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // Проверяем, что изменение произошло в листе 'current' и в ячейке H2
  if (sheet.getName() === 'current' && range.getA1Notation() === 'H2') {
    const newValue = e.value;
    if (newValue === 'Убрать в архив (автоматизация)') {
      archiveData();
    }
  }
}

/**
 * Функция для архивации данных.
 * Копирует указанные колонки из листа 'current' в лист 'archive' и очищает их.
 */
function archiveData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName('current');
  const archiveSheet = ss.getSheetByName('archive');

  // Проверяем, существуют ли оба листа
  if (!currentSheet || !archiveSheet) {
    SpreadsheetApp.getUi().alert('Не найден лист "current" или "archive".');
    return;
  }

  // Определяем столбцы для копирования и их целевые начальные столбцы в архиве
  const columnMappings = {
    'A': 'F',
    'B': 'A',
    'C': 'B',
    'D': 'C',
    'E': 'D',
    'F': 'E',
    'L': 'L',
    'M': 'M',
    'N': 'N',
  };

  // Получаем последний заполненный ряд в листе 'current'
  const lastRow = currentSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Нет данных для архивации.');
    return;
  }

  // Функция для преобразования буквенного обозначения колонки в числовое
  const columnToIndex = (col) => {
    let index = 0;
    for (let i = 0; i < col.length; i++) {
      index *= 26;
      index += col.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
    }
    return index;
  };

  // Проходим по каждой паре источник-назначение
  for (const [sourceCol, targetStartCol] of Object.entries(columnMappings)) {
    const sourceColIndex = columnToIndex(sourceCol);
    const targetStartColIndex = columnToIndex(targetStartCol);

    // Получаем данные из текущей колонки, начиная с 2 строки
    const sourceRange = currentSheet.getRange(2, sourceColIndex, lastRow - 1, 1);
    const values = sourceRange.getValues().filter(row => row[0] !== '');

    if (values.length === 0) {
      continue; // Пропускаем, если нет данных для копирования
    }

    // Находим первый пустой столбец в листе 'archive' начиная с targetStartCol
    const archiveLastCol = archiveSheet.getLastColumn();
    let targetColIndex = targetStartColIndex;

    // Проверяем, занят ли целевой столбец
    while (archiveSheet.getRange(1, targetColIndex).getValue() !== '') {
      targetColIndex++;
    }

    // Находим первую пустую строку в целевом столбце
    let archiveTargetRow = 2; // Начинаем с 2 строки, предполагая, что 1 строка - заголовок
    const archiveColumnData = archiveSheet.getRange(2, targetColIndex, archiveSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < archiveColumnData.length; i++) {
      if (archiveColumnData[i][0] === '') {
        archiveTargetRow = i + 2;
        break;
      }
    }

    // Копируем данные в архив
    archiveSheet.getRange(archiveTargetRow, targetColIndex, values.length, 1).setValues(values);

    // Очищаем скопированные данные в листе 'current'
    currentSheet.getRange(2, sourceColIndex, lastRow - 1, 1).clearContent();
  }

  // Сбрасываем значение в H2 после архивации
  currentSheet.getRange('H2').setValue('');

  SpreadsheetApp.getUi().alert('Архивация завершена успешно.');
}
