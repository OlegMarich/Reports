const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// 📅 Сьогоднішня дата
const today = new Date();
const dateIso = today.toISOString().slice(0, 10);

// 🛣️ Шляхи
const templatePath = path.join(__dirname, 'Loading for day.xlsx');
const jsonPath = path.join(__dirname, 'output', dateIso, 'data.json');
const outputPath = path.join(__dirname, 'output', dateIso, 'Loading Completed.xlsx');

// 🔄 Конвертація часу з float → hh:mm
function convertExcelTime(timeFloat) {
  if (typeof timeFloat !== 'number') return '';
  const totalMinutes = Math.round(timeFloat * 24 * 60);
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
}

// 📥 Зчитування даних
if (!fs.existsSync(jsonPath)) {
  console.error(`❌ Не знайдено: ${jsonPath}`);
  process.exit(1);
}
const data = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));

// 🧠 Сортування
data.sort((a, b) => a['Godzina'] - b['Godzina']);

(async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);
  const sheet = workbook.getWorksheet(1); // Перший аркуш

  let currentRow = 2; // Починаємо після заголовків

  const borderStyle = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
  };

  for (const entry of data) {
    const truck = entry['Kierowca'] || '';
    const driver = entry['Driver'] || '';
    const client = entry['Odbiorca'] || '';
    const trailer = truck.split(' ')[1] || '';
    const truckPlate = truck.split(' ')[0] || '';
    const loadingTime = entry['Godzina'] || '';
    const timeWindow = entry['Timewindow start'] || '';
    const qty = entry['Ilość razem'] || '';
    const pal = entry['Pal'] || '';

    const row = sheet.getRow(currentRow);

    const cells = [
      { col: 'A', value: client },
      { col: 'B', value: truckPlate },
      { col: 'C', value: trailer },
      { col: 'D', value: driver },
      { col: 'E', value: loadingTime },
      { col: 'F', value: timeWindow },
      { col: 'G', value: qty },
      { col: 'H', value: pal },
    ];

    for (const { col, value } of cells) {
      const cell = row.getCell(col);
      if (value !== undefined && value !== null && value !== '') {
        cell.value = value;
        cell.border = borderStyle;
      }
    }

    row.commit();
    currentRow++;
  }

  // 🟦 Обведення діапазону A2:Hn навіть якщо деякі клітинки порожні
  const startRow = 2;
  const endRow = startRow + data.length - 1;
  for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
    const row = sheet.getRow(rowNum);
    for (let col = 1; col <= 9; col++) { // A (1) до H (8)
      const cell = row.getCell(col);
      if (!cell.border || !cell.border.top) {
        cell.border = borderStyle;
      }
    }
    row.commit();
  }

  await workbook.xlsx.writeFile(outputPath);
  console.log(`✅ Завершено: файл збережено у ${outputPath}`);
})();
