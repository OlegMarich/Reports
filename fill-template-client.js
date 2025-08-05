const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// ✅ Отримання дати з аргументу
const selectedDate = process.argv[2];
if (!selectedDate) {
  console.error('❌ Не передано дату як аргумент');
  process.exit(1);
}

// 📥 Читання файлу data.json
const jsonPath = path.join(__dirname, 'output', selectedDate, 'data.json');
if (!fs.existsSync(jsonPath)) {
  console.error(`❌ Не знайдено файл data.json для дати ${selectedDate}`);
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
const outputDir = path.join(__dirname, 'output', selectedDate); // ← ДОДАЙ ЦЕ
// 📄 Шлях до шаблону
const templatePath = path.join(__dirname, 'template.xlsx');

// 📦 Логіка обчислення палет
function getBoxesPerPallet(clientName) {
  const name = clientName.toLowerCase();

  const rules = {
    'aldi': 28,
    'lidl': 48,
    'biedronka': 28,
    'spar hrvatska': 48,
    'spar ljubljana': 48,
    'spar': 32,
    'penny': 32,
    'metro': 28,
    'ta-moro': 48,
    'cba': 48,
    'lunnys': 48,
  };

  let boxesPerPallet = 1; // значення за замовчуванням

  for (const [key, value] of Object.entries(rules)) {
    if (name.includes(key)) {
      boxesPerPallet = value;
      break;
    }
  }

  // Якщо лише 1 ящик на палеті — додаємо ще 1 палету
  if (boxesPerPallet === 1) {
    return boxesPerPallet + 1;
  }

  return boxesPerPallet;
}

// 🧠 Групуємо записи по клієнтах
const grouped = {};
data.forEach((entry) => {
  const client = entry['Odbiorca'];
  if (!grouped[client]) grouped[client] = [];
  grouped[client].push(entry);
});

// 🧾 Генерація шаблонів
async function fillTemplate() {
  for (let i = 0; i < data.length; i++) {
    const entry = data[i];
    const client = entry['Odbiorca'];
    const produkt = (entry['Produkt'] || '').toLowerCase();
    const typ = (entry['Typ'] || '').toLowerCase();

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    const mainSheet = workbook.getWorksheet('RAPORT WYDANIA F-NR 15');
    const bioSheet = workbook.worksheets.find((ws) => ws.name.toLowerCase().includes('bio'));

    if (!mainSheet) {
      console.error(`❌ Не знайдено аркуш RAPORT WYDANIA F-NR 15 для ${client}`);
      continue;
    }

    const qty = Number(entry['Ilość razem'] || 0);
    const pal = Number(entry['Pal'] || 0) || Math.ceil(qty / getBoxesPerPallet(client));

    // Основний аркуш
    mainSheet.getCell('J8').value = entry['Data wysyłki'] || '';
    mainSheet.getCell('C8').value = client || '';
    mainSheet.getCell('J25').value = `${qty} (${pal})`;
    mainSheet.getCell('J29').value = entry['Kierowca'] || '';
    mainSheet.getCell('E10').value = entry['Godzina'] || '';

    // Якщо запис стосується біо-бананів — заповнюємо BIO-аркуш
    const isBioBanana = produkt.includes('bio') || typ.includes('bio');
    if (isBioBanana && bioSheet) {
      bioSheet.getCell('J62').value = entry['Data wysyłki'] || '';
      bioSheet.getCell('C62').value = client + ' (BIO)' || '';
      bioSheet.getCell('J71').value = `${qty} (${pal})`;
      bioSheet.getCell('K65').value = entry['Kierowca'] || '';
      bioSheet.getCell('E63').value = entry['Godzina'] || '';
    }

    const safeClientName = client.replace(/[\\/:*?"<>|]/g, '_');
    const clientBaseDir = path.join(outputDir, safeClientName);
    if (!fs.existsSync(clientBaseDir)) fs.mkdirSync(clientBaseDir, {recursive: true});

    const fileName = `Quality report ${safeClientName}_${i + 1}.xlsx`;
    const outputPath = path.join(clientBaseDir, fileName);

    await workbook.xlsx.writeFile(outputPath);
    console.log(`📄 Створено файл: ${outputPath}`);
  }

  console.log('✅ Усі звіти згенеровано успішно!');
}

fillTemplate().catch(console.error);
