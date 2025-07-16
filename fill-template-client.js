const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// ✅ Отримання дати з аргументу
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
const templatePath = path.join(__dirname, 'template.xlsx');

// 📦 Логіка обчислення палет
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

  // Якщо лише 1 ящик на палеті — додаємо ще 1 палету
  if (boxesPerPallet === 1) {
    return boxesPerPallet + 1;
  }

  return boxesPerPallet;
}

// 🧠 Групуємо записи по клієнтах
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
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    const mainSheet = workbook.getWorksheet('RAPORT WYDANIA F-NR 15');
    const secondSheet = workbook.worksheets[1]; // або workbook.getWorksheet('BIO')

    if (!mainSheet) {
      console.error(`❌ Не знайдено аркуш RAPORT WYDANIA F-NR 15 для ${client}`);
      continue;
    }

    // 🔹 Перший запис (звичайний)
    const entry1 = entries[0];
    const qty1 = Number(entry1['Ilość razem'] || 0);
    const pal1 = Number(entry1['Pal'] || 0) || Math.ceil(qty1 / getBoxesPerPallet(client));

    mainSheet.getCell('J8').value = entry1['Data wysyłki'] || '';
    mainSheet.getCell('C8').value = client || '';
    mainSheet.getCell('J25').value = `${qty1} (${pal1})`;
    mainSheet.getCell('J29').value = entry1['Kierowca'] || '';
    mainSheet.getCell('E10').value = entry1['Godzina'] || '';

    // 🔸 Другий запис (наприклад, BIO) — якщо є
    if (entries.length > 1 && secondSheet) {
      const entry2 = entries[1];
      const qty2 = Number(entry2['Ilość razem'] || 0);
      const pal2 = Number(entry2['Pal'] || 0) || Math.ceil(qty2 / getBoxesPerPallet(client));

      secondSheet.getCell('J62').value = entry2['Data wysyłki'] || '';
      secondSheet.getCell('C62').value = client + ' (BIO)' || '';
      secondSheet.getCell('J71').value = `${qty2} (${pal2})`;
      secondSheet.getCell('K65').value = entry2['Kierowca'] || '';
      secondSheet.getCell('E63').value = entry2['Godzina'] || '';
    }

    const safeClientName = client.replace(/[\\/:*?"<>|]/g, '_');
    const fileName = `${safeClientName}.xlsx`;
    const outputPath = path.join(outputDir, fileName);

    await workbook.xlsx.writeFile(outputPath);
    console.log(`📄 Створено файл: ${outputPath}`);
  }

  console.log('✅ Усі звіти згенеровано успішно!');
}

fillTemplate().catch(console.error);
