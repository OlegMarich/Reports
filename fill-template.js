const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// 📅 Сьогоднішня дата
const today = new Date();
const currentDay = String(today.getDate()).padStart(2, '0');
const currentMonth = String(today.getMonth() + 1).padStart(2, '0');
const dateIso = today.toISOString().slice(0, 10); // YYYY-MM-DD

// 🧾 Шляхи
const templatePath = path.join(__dirname, 'template.xlsx');
const jsonPath = path.join(__dirname, 'output', dateIso, 'data.json'); // ❗️зчитує з output/дата/data.json
const outputDir = path.join(__dirname, 'output', dateIso);

// 📦 Перевірка наявності шаблону і JSON
if (!fs.existsSync(jsonPath)) {
  console.error(`❌ Файл ${jsonPath} не знайдено. Спочатку згенеруй його!`);
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
const workbook = new ExcelJS.Workbook();

async function fillTemplate() {
  await workbook.xlsx.readFile(templatePath);

  const sheet = workbook.getWorksheet('RAPORT WYDANIA F-NR 15');
  if (!sheet) {
    console.error('❌ Аркуш "RAPORT WYDANIA F-NR 15" не знайдено в шаблоні!');
    return;
  }

  for (const entry of data) {
    const newWorkbook = new ExcelJS.Workbook();
    await newWorkbook.xlsx.readFile(templatePath);
    const newSheet = newWorkbook.getWorksheet('RAPORT WYDANIA F-NR 15');

    newSheet.getCell('J8').value = entry['Data wysyłki'];
    newSheet.getCell('C8').value = entry['Odbiorca'];
    newSheet.getCell('J25').value = entry['Ilość разом'];
    newSheet.getCell('J29').value = entry['Kierowca'];

    const safeClientName = entry['Odbiorca'].replace(/[\\/:*?"<>|]/g, '_');
    const outputPath = path.join(outputDir, `${safeClientName}.xlsx`);

    await newWorkbook.xlsx.writeFile(outputPath);
    console.log(`📄 Створено файл: ${outputPath}`);
  }

  console.log('✅ Усі звіти згенеровано успішно!');
}

fillTemplate().catch(console.error);
