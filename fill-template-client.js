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
const outputDir = path.join(__dirname, 'output', selectedDate);
const templatePath = path.join(__dirname, 'template.xlsx');

// 📦 Логіка обчислення палет
function getBoxesPerPallet(clientName) {
  const name = (clientName || '').toLowerCase();
  const rules = {
    'aldi': 28, 'lidl': 48, 'biedronka': 28, 'spar hrvatska': 48,
    'spar ljubljana': 48, 'spar ullo': 32, 'spar bicske': 32, 'penny': 32, 'metro': 28,
    'ta-moro': 48, 'cba': 48, 'lunnys': 48, 'horti': 54,
  };
  let boxesPerPallet = 1;
  for (const [key, value] of Object.entries(rules)) {
    if (name.includes(key)) {
      boxesPerPallet = value;
      break;
    }
  }
  return boxesPerPallet === 1 ? 2 : boxesPerPallet;
}

// 🧠 Надійне визначення BIO
function isBioEntry(entry) {
  const odb = (entry['Odbiorca'] || '').toLowerCase();
  const produkt = (entry['Produkt'] || '').toLowerCase();
  const typ = (entry['Typ'] || '').toLowerCase();
  const line = (entry['Linia'] || entry['Line'] || entry['Nazwa linii'] || '').toLowerCase();
  const re = /\bbio\b/;
  return re.test(odb) || re.test(produkt) || re.test(typ) || re.test(line);
}

// 🧾 Генерація шаблонів для кожного запису
async function fillTemplate() {
  let idx = 1;
  for (const entry of data) {
    const client = (entry['Odbiorca'] || '').replace(/\s*\(.*bio.*\)/i, '').trim();
    const truck = entry['Kierowca'] || 'unknown';
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    const mainSheet = workbook.getWorksheet('RAPORT WYDANIA F-NR 15');
    if (!mainSheet) {
      console.error('❌ Не знайдено аркуш "RAPORT WYDANIA F-NR 15" у шаблоні');
      continue;
    }

    const qty = Number(entry['Ilość razem'] || 0);
    const palGiven = Number(entry['Pal'] || 0);
    const isBio = isBioEntry(entry);
    const pal = palGiven > 0 ? palGiven : (qty > 0 ? Math.ceil(qty / getBoxesPerPallet(client)) : 0);

    if (isBio) {
      // BIO блок (нижній)
      mainSheet.getCell('J60').value = entry['Data wysyłki'] || '';
      mainSheet.getCell('C60').value = `${client} (BIO)`;
      mainSheet.getCell('J69').value = `${qty} (${pal})`;
      mainSheet.getCell('K63').value = entry['Kierowca'] || '';
      mainSheet.getCell('E61').value = entry['Godzina'] || '';
    } else {
      // Банани (верхній блок)
      mainSheet.getCell('J8').value = entry['Data wysyłki'] || '';
      mainSheet.getCell('C8').value = client || '';
      mainSheet.getCell('J25').value = `${qty} (${pal})`;
      mainSheet.getCell('J29').value = entry['Kierowca'] || '';
      mainSheet.getCell('E10').value = entry['Godzina'] || '';
    }

    // Лог для контролю
    console.log(`➡️  ${client} [${truck}]  banana=${!isBio ? qty : 0} / bio=${isBio ? qty : 0}`);

    // 📂 Збереження файлу
    const safeClientName = client.replace(/[\\/:*?"<>|]/g, '_');
    const safeTruck = truck.replace(/[\\/:*?"<>|]/g, '_');
    const clientBaseDir = path.join(outputDir, safeClientName);
    if (!fs.existsSync(clientBaseDir)) fs.mkdirSync(clientBaseDir, { recursive: true });

    // Додаємо унікальний індекс або час, щоб не було перезапису
    const uniqueId = entry['Godzina'] ? entry['Godzina'].replace(/[: ]/g, '-') : idx;
    const fileName = `Quality report ${safeClientName}_${safeTruck}_${uniqueId}.xlsx`;
    const outputPath = path.join(clientBaseDir, fileName);

    await workbook.xlsx.writeFile(outputPath);
    console.log(`📄 Створено файл: ${outputPath}`);
    idx++;
  }

  console.log('✅ Усі звіти згенеровано успішно!');
}

fillTemplate().catch(err => {
  console.error('❌ Помилка:', err);
  process.exit(1);
});
