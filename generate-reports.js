const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');

const inputDir = path.join(__dirname, 'input');

// ⏱ Отримання дати з аргументу
const userDateArg = process.argv[2]; // очікується у форматі YYYY-MM-DD
const today = new Date();

let targetDate;
let date;

if (userDateArg && /^\d{4}-\d{2}-\d{2}$/.test(userDateArg)) {
  const [year, month, day] = userDateArg.split('-');
  targetDate = `${day}.${month}`;
  date = userDateArg;
} else {
  const currentDay = String(today.getDate()).padStart(2, '0');
  const currentMonth = String(today.getMonth() + 1).padStart(2, '0');
  targetDate = `${currentDay}.${currentMonth}`;
  date = today.toISOString().slice(0, 10);
}

// 🔍 Зчитування файлів
const files = fs.readdirSync(inputDir);
const transportPlanFile = files.find(f => f.toLowerCase().includes('plan_week'));
const salesPlanFile = files.find(f => f.toLowerCase().includes('sales plan'));

if (!transportPlanFile || !salesPlanFile) {
  console.error('❌ Файли не знайдено.');
  process.exit(1);
}

const transportPath = path.join(inputDir, transportPlanFile);
const salesPath = path.join(inputDir, salesPlanFile);

if (fs.statSync(transportPath).size === 0) {
  console.error(`❌ Файл ${transportPlanFile} порожній або пошкоджений.`);
  process.exit(1);
}
if (fs.statSync(salesPath).size === 0) {
  console.error(`❌ Файл ${salesPlanFile} порожній або пошкоджений.`);
  process.exit(1);
}

const transportWorkbook = xlsx.readFile(transportPath);
const salesWorkbook = xlsx.readFile(salesPath);

// 🧠 Пошук аркуша за датою
function normalizeDateString(str) {
  return str.replace(/\D/g, '').padStart(4, '0');
}

function findSheetByDate(sheetNames, ddmm) {
  const normalizedTarget = normalizeDateString(ddmm);
  return sheetNames.find(name => normalizeDateString(name).includes(normalizedTarget));
}

const matchedSheetName = findSheetByDate(transportWorkbook.SheetNames, targetDate);
if (!matchedSheetName) {
  console.error(`❌ Не знайдено аркуша з назвою ${targetDate}`);
  process.exit(1);
}

const transportSheet = transportWorkbook.Sheets[matchedSheetName];
const salesSheet = salesWorkbook.Sheets[salesWorkbook.SheetNames[0]];

const transportData = xlsx.utils.sheet_to_json(transportSheet, { defval: '', range: 0 });
const salesData = xlsx.utils.sheet_to_json(salesSheet, { defval: '' });

function normalizeRow(row) {
  const normalized = {};
  for (const key in row) {
    normalized[key.toLowerCase().trim()] = row[key];
  }
  return normalized;
}

function convertExcelTime(excelTime) {
  if (isNaN(excelTime)) return '';
  const totalMinutes = Math.round(excelTime * 24 * 60);
  const hours = String(Math.floor(totalMinutes / 60)).padStart(2, '0');
  const minutes = String(totalMinutes % 60).padStart(2, '0');
  return `${hours}:${minutes}`;
}

transportData.sort((a, b) => {
  const rA = normalizeRow(a);
  const rB = normalizeRow(b);
  return (rA['loading time'] || 0) - (rB['loading time'] || 0);
});

const result = [];
const aldiRows = [];

transportData.forEach(row => {
  const r = normalizeRow(row);
  const client = r['customer'] || '';
  const qty = Number(r['qty']) || 0;
  const pal = Math.ceil(Number(r['pal']) || 0); // ✅ округлення палет
  const truck = `${r['truck plate nr'] || ''} ${r['trailer plate nr'] || ''}`.trim();
  const driver = r['driver'] || '';
  const loading = convertExcelTime(Number(r['loading time']));
  const start = convertExcelTime(Number(r['timewindow start']));

  if (!client) return;

  if (client.toLowerCase().includes('aldi') && client.toLowerCase().includes('lukovica')) {
    aldiRows.push({ qty, pal, driver, loading, start, truck });
  } else {
    result.push({
      'Data wysyłki': date,
      'Odbiorca': client,
      'Ilość razem': qty,
      'Kierowca': truck,
      'Pal': pal,
      'Driver': driver,
      'Godzina': loading,
      'Timewindow start': start,
    });
  }
});

if (aldiRows.length > 0) {
  const totalQty = aldiRows.reduce((sum, r) => sum + r.qty, 0);
  const totalPal = Math.ceil(aldiRows.reduce((sum, r) => sum + r.pal, 0)); // ✅ округлення підсумку
  const lastWithTruck = [...aldiRows].reverse().find(r => r.truck);
result.push({
  'Data wysyłki': date,
  'Odbiorca': 'Aldi Lukovica',
  'Ilość razem': totalQty,
  'Kierowca': lastWithTruck?.truck || '',
  'Pal': totalPal,
  'Driver': lastWithTruck?.driver || '',
  'Godzina': lastWithTruck?.loading || '',
  'Timewindow start': lastWithTruck?.start || '',
});

}

const outputPath = path.join(__dirname, 'output', date);
if (!fs.existsSync(outputPath)) fs.mkdirSync(outputPath, { recursive: true });

fs.writeFileSync(path.join(outputPath, 'data.json'), JSON.stringify(result, null, 2), 'utf-8');
console.log(`✅ Збережено у: ${path.join(outputPath, 'data.json')}`);
