const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

// 🧽 Очищення зайвих суфіксів (з базовою нормалізацією)
function normalizeName(name) {
  return name
    .toLowerCase()
    .replace(/\[.*?\]/g, '') // видалити [3+/4] або [4]
    .replace(/\(.*?\)/g, '') // видалити (3-3.5)
    //.replace(/\bbio\b/gi, '') // видалити "bio"
    //node parser-sales.js
.replace(/\bbananas?\b/gi, '') // видалити "banana" або "bananas"
    .replace(/\s+/g, ' ') // прибрати подвійні пробіли
    .trim();
}

// Функція капіталізації кожного слова в рядку
function capitalizeWords(str) {
  return str.replace(/\b\w/g, (char) => char.toUpperCase());
}

// 📅 Назва дня тижня англійською
function getEnglishWeekdayName(dateStr) {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const date = new Date(dateStr);
  if (isNaN(date)) {
    console.warn(`⚠️ Некоректна дата: "${dateStr}"`);
    return null;
  }
  return days[date.getDay()];
}

function askWeek() {
  return new Promise((resolve) => {
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });
    rl.question('Введіть номер тижня (наприклад, 26): ', (answer) => {
      rl.close();
      resolve(answer.trim());
    });
  });
}

(async () => {
  let week = process.argv[2];
  if (!week) {
    week = await askWeek();
  }
  if (!week) {
    console.error('❌ Не передано номер тижня');
    process.exit(1);
  }

  const folderName = `${week}_Week`;

  const jsonPath = path.join(__dirname, 'output', folderName, 'sales.json');
  const templatePath = path.join(__dirname, 'template-week-plan.xlsx');
  const outputPath = path.join(__dirname, 'output', folderName, `PLAN_week_${week}.xlsx`);

  if (!fs.existsSync(jsonPath)) {
    console.error(`❌ Не знайдено файл: ${jsonPath}`);
    process.exit(1);
  }
  if (!fs.existsSync(templatePath)) {
    console.error(`❌ Не знайдено шаблон: ${templatePath}`);
    process.exit(1);
  }

  const salesData = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));

  const workbook = new ExcelJS.Workbook();

  await workbook.xlsx.readFile(templatePath);

  console.log('📋 Назви листів у шаблоні:');
  workbook.worksheets.forEach((ws) => console.log(`- "${ws.name}"`));

  function findSheetIgnoreCase(workbook, name) {
    const lowerName = name.trim().toLowerCase();
    return workbook.worksheets.find((sheet) => sheet.name.trim().toLowerCase() === lowerName);
  }

  function rowExists(sheet, customer, qty, product) {
    for (let i = 2; i <= sheet.actualRowCount; i++) {
      const row = sheet.getRow(i);
      if (
        normalizeName(row.getCell(2).value || '') === normalizeName(customer) &&
        row.getCell(12).value === qty &&
        normalizeName(row.getCell(11).value || '') === normalizeName(product) &&
        (row.getCell(1).value || '') === 'OUTBOUND'
      ) {
        return true;
      }
    }
    return false;
  }

  for (const client of salesData) {
    if (normalizeName(client.customer) === 'total') continue;

    for (const day of client.data) {
      const { date, qty } = day;
      if (!qty || qty === 0) continue;

      const weekdayName = getEnglishWeekdayName(date);
      if (!weekdayName) continue;

      const sheet = findSheetIgnoreCase(workbook, weekdayName);
      if (!sheet) {
        console.warn(`⚠️ Не знайдено вкладки "${weekdayName}", пропускаю...`);
        continue;
      }

      const baseCustomer = normalizeName(client.customer);
      const unloadingPlace = normalizeName(client.line || '');
      const fullCustomer = capitalizeWords(`${baseCustomer} ${unloadingPlace}`.trim());

      const product = normalizeName(client.line).includes('bio') ? 'BIO BANANA' : 'BANANA';

      const safeBoxPerPal = 24;

      const weightPerBox = 0;
      const grossWeight = qty * weightPerBox;

      const fullPallets = Math.floor(qty / safeBoxPerPal);
      const hasRemainder = qty % safeBoxPerPal > 0;
      const pal = fullPallets + (hasRemainder ? 1 : 0);

      if (rowExists(sheet, fullCustomer, qty, product)) continue;

      const rowIndex = sheet.actualRowCount + 1;
      const row = sheet.getRow(rowIndex);

      row.getCell(1).value = 'OUTBOUND';
      row.getCell(2).value = fullCustomer;
      row.getCell(3).value = '';
      row.getCell(4).value = '';
      row.getCell(5).value = '';
      row.getCell(6).value = '';
      row.getCell(7).value = '';
      row.getCell(8).value = 'Nagytarcsa';
      row.getCell(9).value = '';
      row.getCell(10).value = '';
      row.getCell(11).value = product;
      row.getCell(12).value = qty;
      row.getCell(13).value = pal;
      row.getCell(14).value = '';
      row.getCell(15).value = '';
      row.getCell(16).value = grossWeight;
      row.getCell(17).value = '';
      row.getCell(18).value = '';
      row.getCell(19).value = '';
      row.getCell(20).value = '';
      row.getCell(21).value = '';

      row.commit();
    }
  }

  await workbook.xlsx.writeFile(outputPath);
  console.log(`✅ План успішно збережено: ${outputPath}`);
})();
