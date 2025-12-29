const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// ✅ Дата
const selectedDate = process.argv[2];
if (!selectedDate) {
  console.error('❌ Не передано дату. Приклад: node generate.js 2025-12-01');
  process.exit(1);
}

// 📥 Шляхи
const baseDir = __dirname;
const outputDir = path.join(baseDir, 'output', selectedDate);
const jsonPath = path.join(outputDir, 'data.json');
const templatePath = path.join(baseDir, 'shipping card.xlsx');

// 🔎 Перевірки
if (!fs.existsSync(jsonPath)) {
  console.error(`❌ Не знайдено data.json: ${jsonPath}`);
  process.exit(1);
}
if (!fs.existsSync(templatePath)) {
  console.error(`❌ Не знайдено shipping card.xlsx`);
  process.exit(1);
}

// 📥 Дані
let data;
try {
  data = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
  if (!Array.isArray(data)) throw new Error('data.json не масив');
} catch (e) {
  console.error('❌ Помилка JSON:', e);
  process.exit(1);
}

// 🔢 Helpers
const parseQty = (v) => Number(typeof v === 'string' ? v.replace(',', '.').trim() : v) || 0;

const safeName = (s) =>
  String(s || '')
    .replace(/[\\/:*?"<>|]/g, '_')
    .trim() || 'unknown';

const isBioEntry = (e) => {
  const re = /\bbio\b/;
  return ['Odbiorca', 'Produkt', 'Typ', 'Linia', 'Line'].some((k) =>
    re.test(String(e[k] || '').toLowerCase()),
  );
};

// 🧾 Заповнення шаблону
async function createShippingCard(entry, mode, index) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const sheet = workbook.getWorksheet('KARTA');
  if (!sheet) throw new Error('Немає аркуша KARTA');

  const client = (entry['Odbiorca'] || '').replace(/\( *bio[^\)]*\)/i, '').trim();

  const car = entry['Kierowca'] || 'unknown';
  const driver = entry['Driver'] || '';
  const date = entry['Data wysyłki'] || '';
  const palletType = entry['Pallet type'] || '';

  const qty = parseQty(entry['Ilość razem']);
  const pal = parseQty(entry['Pal']);

  // Заголовки
  sheet.getCell('A1').value = 'KARTA WYSYŁKOWA / SHIPPING CARD';
  sheet.getCell('G1').value = `Data/Date: ${date}`;
  sheet.getCell('B11').value = `DRIVER: ${driver}`;
  sheet.getCell('B13').value = `CAR NUMBER: ${car}`;
  sheet.getCell('B15').value = `DESTINATION: ${client}`;
  sheet.getCell('H26').value = palletType;
  sheet.getCell('H3').value = qty;

  // BAN / BIO
  if (mode === 'BAN') {
    sheet.getCell('A27').value = 'Banana';
    sheet.getCell('D27').value = qty;
    sheet.getCell('H27').value = pal;
  } else {
    sheet.getCell('A27').value = 'BIO banana';
    sheet.getCell('D27').value = qty;
    sheet.getCell('H27').value = pal;
  }

  // Збереження
  const clientDir = path.join(outputDir, safeName(client));
  if (!fs.existsSync(clientDir)) fs.mkdirSync(clientDir, {recursive: true});

  const fileName = `Shipping card ${index} - ${safeName(client)}_${safeName(car)}_${mode}.xlsx`;

  const filePath = path.join(clientDir, fileName);
  await workbook.xlsx.writeFile(filePath);

  console.log(`📄 Створено: ${filePath}`);
}

// ▶️ Головна логіка
async function run() {
  let index = 0;

  for (const entry of data) {
    const mode = isBioEntry(entry) ? 'BIO' : 'BAN';
    index++;
    await createShippingCard(entry, mode, index);
  }

  console.log(`✅ Генерацію завершено. Файлів: ${index}`);
}


run().catch((err) => {
  console.error('❌ Критична помилка:', err);
  process.exit(1);
});
