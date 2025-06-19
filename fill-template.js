const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// 📅 Сьогоднішня дата
const today = new Date();
const dateIso = today.toISOString().slice(0, 10); // YYYY-MM-DD

// 🧾 Шляхи
const templatePath = path.join(__dirname, 'template.xlsx');
const jsonPath = path.join(__dirname, 'output', dateIso, 'data.json');
const outputDir = path.join(__dirname, 'output', dateIso);

// ❗ Перевірка наявності JSON
if (!fs.existsSync(jsonPath)) {
  console.error(`❌ Файл ${jsonPath} не знайдено. Спочатку згенеруй його!`);
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));

// 🧠 Кількість ящиків на палету для кожного клієнта
function getBoxesPerPallet(clientName) {
  const name = clientName.toLowerCase();

  if (name.includes('aldi')) return 28;
  if (name.includes('lidl')) return 48;
  if (name.includes('biedronka')) return 28;
  if (name.includes('spar hrvatska')) return 48;
  if (name.includes('spar ljubljana')) return 48;
  if (name.includes('spar')) return 32;

  return 1; // За замовчуванням
}

async function fillTemplate() {
  for (const entry of data) {
    const newWorkbook = new ExcelJS.Workbook();
    await newWorkbook.xlsx.readFile(templatePath);
    const sheet = newWorkbook.getWorksheet('RAPORT WYDANIA F-NR 15');

    if (!sheet) {
      console.error('❌ Аркуш "RAPORT WYDANIA F-NR 15" не знайдено в шаблоні!');
      continue;
    }

    const client = entry['Odbiorca'];
    const qty = Number(entry['Ilość razem'] || 0);
    const providedPal = Number(entry['Pal'] || 0);

    // Обчислення палет, якщо не задано
    const pal = providedPal > 0 ? providedPal : Math.ceil(qty / getBoxesPerPallet(client));

    sheet.getCell('J8').value = entry['Data wysyłki'];
    sheet.getCell('C8').value = client;
    sheet.getCell('J25').value = `${qty} (${pal})`;
    sheet.getCell('J29').value = entry['Kierowca'];

    const safeClientName = client.replace(/[\\/:*?"<>|]/g, '_');
    const outputPath = path.join(outputDir, `${safeClientName}.xlsx`);

    await newWorkbook.xlsx.writeFile(outputPath);
    console.log(`📄 Створено файл: ${outputPath}`);
  }

  console.log('✅ Усі звіти згенеровано успішно!');
}

fillTemplate().catch(console.error);
