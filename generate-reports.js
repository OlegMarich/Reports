const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');

// 📁 Шляхи
const inputDir = path.join(__dirname, 'input');

// 🔍 Зчитування назв файлів
const files = fs.readdirSync(inputDir);
const transportPlanFile = files.find((f) => f.toLowerCase().includes('plan_week'));
const salesPlanFile = files.find((f) => f.toLowerCase().includes('sales plan'));

if (!transportPlanFile || !salesPlanFile) {
  console.error('❌ Не знайдено обидва файли: транспортний план або sales plan.');
  process.exit(1);
}

const transportPlanPath = path.join(inputDir, transportPlanFile);
const salesPlanPath = path.join(inputDir, salesPlanFile);

// 📖 Зчитування Excel
const transportWorkbook = xlsx.readFile(transportPlanPath);
const salesWorkbook = xlsx.readFile(salesPlanPath);

// 📆 Параметр дати з командного рядка
const userDateArg = process.argv[2]; // Очікується у форматі DD.MM
const today = new Date();
const currentDay = String(today.getDate()).padStart(2, '0');
const currentMonth = String(today.getMonth() + 1).padStart(2, '0');
const targetDate = userDateArg || `${currentDay}.${currentMonth}`;

// 🧠 Пошук аркуша з відповідною назвою
function findSheetByDate(sheetNames, ddmm) {
  return sheetNames.find((name) => name.startsWith(ddmm));
}

const matchedSheetName = findSheetByDate(transportWorkbook.SheetNames, targetDate);

if (!matchedSheetName) {
  console.error(`❌ Не знайдено аркуша з назвою ${targetDate} у транспортному плані!`);
  process.exit(1);
}

// 📅 Формування ISO-дати
function toIsoDate(ddmm) {
  const [day, month] = ddmm.split('.');
  const year = today.getFullYear();
  return new Date(`${year}-${month}-${day}`).toISOString().slice(0, 10);
}

const date = toIsoDate(targetDate);
const transportSheet = transportWorkbook.Sheets[matchedSheetName];
const salesSheet = salesWorkbook.Sheets[salesWorkbook.SheetNames[0]];

// 🔄 Конвертація аркушів у JSON
const transportData = xlsx.utils.sheet_to_json(transportSheet, { defval: '', range: 0 });
const salesData = xlsx.utils.sheet_to_json(salesSheet, { defval: '' });

// 🔧 Нормалізація ключів
function normalizeRow(row) {
  const normalized = {};
  for (const key in row) {
    normalized[key.toLowerCase().trim()] = row[key];
  }
  return normalized;
}

// 📦 Формування результату
const result = [];
const aldiRows = [];

transportData.forEach((row) => {
  const r = normalizeRow(row);
  const client = r['customer'] || '';
  const quantity = Number(r['qty']);
  const pallets = Number(r['pal']);
  const truck = `${r['truck plate nr']} ${r['trailer plate nr'] || ''}`.trim();

  if (!client) return;

  if (client.toLowerCase().includes('aldi') && client.toLowerCase().includes('lukovica')) {
    aldiRows.push({ quantity, pallets });
  } else {
    result.push({
      'Data wysyłki': date,
      'Odbiorca': client,
      'Ilość razem': quantity,
      'Kierowca': truck,
      'Pal': pallets,
    });
  }
});

if (aldiRows.length > 0) {
  const totalQty = aldiRows.reduce((sum, r) => sum + r.quantity, 0);
  const totalPal = aldiRows.reduce((sum, r) => sum + r.pallets, 0);
  result.push({
    'Data wysyłki': date,
    'Odbiorca': 'Aldi Lukovica',
    'Ilość razem': totalQty,
    'Kierowca': '',
    'Pal': totalPal,
  });
}

// 📁 Створення папки з назвою дати
const outputDir = path.join(__dirname, 'output', date);
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

// 💾 Запис у файл
const outputPath = path.join(outputDir, 'data.json');
fs.writeFileSync(outputPath, JSON.stringify(result, null, 2), 'utf-8');
console.log(`✅ Звіт за ${date} збережено у: ${outputPath}`);
