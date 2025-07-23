const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// ✅ 1. Отримання дати
const selectedDate = process.argv[2];
if (!selectedDate) {
  console.error('❌ Не передано дату як аргумент');
  process.exit(1);
}

const jsonPath = path.join(__dirname, 'output', selectedDate, 'data.json');
if (!fs.existsSync(jsonPath)) {
  console.error(`❌ Не знайдено файл data.json для дати ${selectedDate}`);
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
const outputDir = path.join(__dirname, 'output', selectedDate);

// 📄 Шлях до шаблону
const templatePath = path.join(__dirname, 'shiping card.xlsx');

// 📦 Обчислення кількості палет
function getBoxesPerPallet(clientName) {
  const name = clientName.toLowerCase();
  let boxesPerPallet = 1;

  if (name.includes('aldi')) boxesPerPallet = 28;
  else if (name.includes('lidl')) boxesPerPallet = 48;
  else if (name.includes('biedronka')) boxesPerPallet = 28;
  else if (name.includes('spar hrvatska')) boxesPerPallet = 48;
  else if (name.includes('spar ljubljana')) boxesPerPallet = 48;
  else if (name.includes('spar')) boxesPerPallet = 32;
  else if (name.includes('penny')) boxesPerPallet = 32;
  else if (name.includes('metro')) boxesPerPallet = 28;
  else if (name.includes('ta-moro')) boxesPerPallet = 48;
  else if (name.includes('cba')) boxesPerPallet = 48;
  else if (name.includes('lunnys')) boxesPerPallet = 48;

  if (boxesPerPallet === 1) return 2;
  return boxesPerPallet;
}

// 🧠 Групуємо по клієнтах
const grouped = {};
data.forEach(entry => {
  const client = entry['Odbiorca'];
  if (!grouped[client]) grouped[client] = [];
  grouped[client].push(entry);
});

// 🧾 Генерація шаблонів
async function fillTemplate() {
  for (const client in grouped) {
    const entries = grouped[client];
    const entry = entries[0];

    const qty = Number(entry['Ilość razem'] || 0);
    const pal = Number(entry['Pal'] || 0) || Math.ceil(qty / getBoxesPerPallet(client));

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    const sheet = workbook.getWorksheet('KARTA');
    if (!sheet) {
      console.error(`❌ Не знайдено аркуш "KARTA" для ${client}`);
      continue;
    }

    // 📌 Заповнення клітинок
    sheet.getCell('A1').value = `KARTA WYSYŁKOWA/SHIPPING CARD     Data/Date ${entry['Data wysyłki'] || ''}`;
    sheet.getCell('B11').value = entry['Kierowca'] || '';
    sheet.getCell('B13').value = entry['Nr auta'] || '';
    sheet.getCell('B15').value = client || '';
    sheet.getCell('B20').value = entry['Godzina'] || '';
    sheet.getCell('D26').value = qty;
    sheet.getCell('H26').value = pal;

    const safeClientName = client.replace(/[\\/:*?"<>|]/g, '_');
    const outputPath = path.join(outputDir, `${safeClientName}_card.xlsx`);

    await workbook.xlsx.writeFile(outputPath);
    console.log(`✅ Створено файл: ${outputPath}`);
  }

  console.log('🎉 Усі shipping cards згенеровано!');
}

fillTemplate().catch(console.error);
