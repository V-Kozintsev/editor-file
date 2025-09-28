import XLSX from 'xlsx';

const INPUT_FILE_NAME = 'отч.xls';

// Читаем Excel-книгу и первый лист
const workbook = XLSX.readFile(INPUT_FILE_NAME);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Получаем все данные листа как массив массивов (без заголовков)
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

// Ищем строку с "Номенклатура"
const nomenklaturaIndex = data.findIndex(
  row => Array.isArray(row) && row.includes('Номенклатура')
);
if (nomenklaturaIndex === -1) throw new Error('Строка с "Номенклатура" не найдена');
const nomenklaturaRow = data[nomenklaturaIndex];

// Ищем строку с "Количество" и "Выручка"
const colsIndex = data.findIndex(
  row => Array.isArray(row) && row.includes('Количество') && row.includes('Выручка')
);
if (colsIndex === -1) throw new Error('Строка с "Количество" и "Выручка" не найдена');
const colsRow = data[colsIndex];

// Определяем индексы искомых колонок
const idxNomenklatura = nomenklaturaRow.indexOf('Номенклатура');
const idxKol = colsRow.indexOf('Количество');
const idxVir = colsRow.indexOf('Выручка');

console.log('Индексы колонок:', { idxNomenklatura, idxKol, idxVir });

// Начинаем читать данные после последней из найденных строк с заголовками
const startDataIndex = Math.max(nomenklaturaIndex, colsIndex) + 1;
const dataRows = data.slice(startDataIndex);

// Фильтруем и формируем массив объектов с нужными колонками
const filteredData = dataRows
  .filter(row => row && row[idxNomenklatura]) // "Номенклатура" обязательна
  .map(row => ({
    Номенклатура: row[idxNomenklatura] || '',
    Количество: row[idxKol] || '',
    Выручка: row[idxVir] || ''
  }));

// Подсчет сумм (с учетом чисел с запятой)
const sumKol = filteredData.reduce((acc, cur) => {
  const val = parseFloat(String(cur.Количество).replace(',', '.'));
  return acc + (isNaN(val) ? 0 : val);
}, 0);

const sumVir = filteredData.reduce((acc, cur) => {
  const val = parseFloat(String(cur.Выручка).replace(',', '.'));
  return acc + (isNaN(val) ? 0 : val);
}, 0);

// Добавляем строку с итогами
filteredData.push({
  Номенклатура: 'Итого',
  Количество: sumKol,
  Выручка: sumVir
});

// Записываем в новый Excel
const newWorkbook = XLSX.utils.book_new();
const newWorksheet = XLSX.utils.json_to_sheet(filteredData);
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'filtered');
XLSX.writeFile(newWorkbook, 'filtered_result.xlsx');

console.log('Файл filtered_result.xlsx создан');
