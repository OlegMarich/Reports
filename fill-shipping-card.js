const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// ✅ Отримання дати
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
const templatePath = path.join(__dirname, 'shipping card.xlsx');

function parseQty(value) {
  if (typeof value === 'string') {
    value = value.replace(',', '.').trim();
  }
  return Number(value) || 0;
}

// Нормалізуємо клієнта, прибираючи “(Bio bananas)” і подібне
function canonicalClientName(name) {
  if (!name) return '';
  return name
    .replace(/\( *bio[^\)]*\)/i, '') // прибрати (Bio ...)
    .replace(/\( *\)/, '') // випадкові пусті дужки
    .trim();
}

// Групування по: канонічний клієнт + авто + дата
function groupByMultipleOrders(data) {
  const grouped = {};
  data.forEach(entry => {
    const clientRaw = entry['Odbiorca'];
    const client = canonicalClientName(clientRaw);
    const car = entry['Kierowca'];
    const date = entry['Data wysyłki'];
    const key = `${client}__${car}__${date}`;

    if (!grouped[key]) grouped[key] = { entries: [], clientCanonical: client, clientRawList: new Set() };
    grouped[key].entries.push(entry);
    grouped[key].clientRawList.add(clientRaw);
  });
  return grouped;
}

const groupedOrders = groupByMultipleOrders(data);

async function fillTemplate() {
  for (const key in groupedOrders) {
    const { entries, clientCanonical, clientRawList } = groupedOrders[key];
    const first = entries[0];

    const clientDisplay = clientCanonical; // об’єднаний ім’я
    const carNumber = first['Kierowca'];
    const driver = first['Driver'] || '';
    const shipDate = first['Data wysyłki'];

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const sheet = workbook.getWorksheet('KARTA');

    if (!sheet) {
      console.error(`❌ Не знайдено аркуш "KARTA"`);
      continue;
    }

    // Заголовки
    sheet.getCell('A1').value = `KARTA WYSYŁKOWA/SHIPPING CARD`;
    sheet.getCell('G1').value = `Data/Date: ${shipDate}`;
    sheet.getCell('B11').value = `DRIVER: ${driver}`;
    sheet.getCell('B13').value = `CAR NUMBER: ${carNumber}`;
    sheet.getCell('B15').value = `DESTINATION: ${clientDisplay}`;

    // Групування по типу товару: стандартні (не bio) і bio
    const totalConvQty = entries
      .filter(e => !((e['Typ'] || '').toString().toLowerCase().includes('bio')))
      .reduce((sum, e) => sum + parseQty(e['Ilość razem']), 0);

    const totalConvPal = entries
      .filter(e => !((e['Typ'] || '').toString().toLowerCase().includes('bio')))
      .reduce((sum, e) => sum + parseQty(e['Pal']), 0);

    const totalBioQty = entries
      .filter(e => ((e['Typ'] || '').toString().toLowerCase().includes('bio')))
      .reduce((sum, e) => sum + parseQty(e['Ilość razem']), 0);

    const totalBioPal = entries
      .filter(e => ((e['Typ'] || '').toString().toLowerCase().includes('bio')))
      .reduce((sum, e) => sum + parseQty(e['Pal']), 0);

    const totalQty = totalConvQty + totalBioQty;

    sheet.getCell('H3').value = totalQty;

    if (totalConvQty > 0) {
      sheet.getCell('A27').value = 'Banana';
      sheet.getCell('D27').value = totalConvQty;
      sheet.getCell('H27').value = totalConvPal;
    }

    if (totalBioQty > 0) {
      sheet.getCell('A28').value = 'BIO banana';
      sheet.getCell('D28').value = totalBioQty;
      sheet.getCell('H28').value = totalBioPal;
    }

    // Збереження, використовуємо канонічне ім'я для папки (але можна включити оригінал)
    const safeClient = clientDisplay.replace(/[\\/:*?"<>|]/g, '_');
    const safeCar = carNumber.replace(/[\\/:*?"<>|]/g, '_');
    const folderPath = path.join(__dirname, 'output', selectedDate, safeClient);

    if (!fs.existsSync(folderPath)) fs.mkdirSync(folderPath, { recursive: true });

    const fileName = `Shipping card ${safeClient} - ${safeCar}.xlsx`;
    const filePath = path.join(folderPath, fileName);

    await workbook.xlsx.writeFile(filePath);
    console.log(`✅ Створено: ${filePath}`);
  }

  console.log('🎉 Всі shipping cards створено!');
}

fillTemplate().catch(console.error);
