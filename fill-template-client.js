
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
    'spar ljubljana': 48, 'spar': 32, 'penny': 32, 'metro': 28,
    'ta-moro': 48, 'cba': 48, 'lunnys': 48,
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

// 🔑 Групування по клієнту + авто
function normalizeClientKey(entry) {
  const client = (entry['Odbiorca'] || '').replace(/\s*\(.*bio.*\)/i, '').trim();
  const truck = entry['Kierowca'] || 'unknown';
  return `${client}__${truck}`;
}

// 🧠 Надійне визначення BIO
function isBioEntry(entry) {
  const odb = (entry['Odbiorca'] || '').toLowerCase();
  const produkt = (entry['Produkt'] || '').toLowerCase();
  const typ = (entry['Typ'] || '').toLowerCase();
  const line = (entry['Linia'] || entry['Line'] || entry['Nazwa linii'] || '').toLowerCase();
  // вважаємо BIO, якщо зустрічається слово "bio" в будь-якому з полів
  // використовуємо \bbio\b щоб уникати випадкових збігів типу "biodegradable"
  const re = /\bbio\b/;
  return re.test(odb) || re.test(produkt) || re.test(typ) || re.test(line);
}

const grouped = {};
for (const entry of data) {
  const key = normalizeClientKey(entry);
  if (!grouped[key]) grouped[key] = [];
  grouped[key].push(entry);
}

// 🧾 Генерація шаблонів
async function fillTemplate() {
  for (const [key, entries] of Object.entries(grouped)) {
    const [client, truck] = key.split('__');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    const mainSheet = workbook.getWorksheet('RAPORT WYDANIA F-NR 15');
    if (!mainSheet) {
      console.error('❌ Не знайдено аркуш "RAPORT WYDANIA F-NR 15" у шаблоні');
      continue;
    }

    // 📊 Агрегуємо банани та біо-банани окремо + зберігаємо перший запис для метаданих
    const totals = { banana: { qty: 0, pal: 0 }, bio: { qty: 0, pal: 0 } };
    let firstBanana = null;
    let firstBio = null;

    for (const entry of entries) {
      const qty = Number(entry['Ilość razem'] || 0);
      const palGiven = Number(entry['Pal'] || 0);
      const isBio = isBioEntry(entry);

      const pal = palGiven > 0 ? palGiven : (qty > 0 ? Math.ceil(qty / getBoxesPerPallet(client)) : 0);

      if (isBio) {
        totals.bio.qty += qty;
        totals.bio.pal += pal;
        if (!firstBio) firstBio = entry;
      } else {
        totals.banana.qty += qty;
        totals.banana.pal += pal;
        if (!firstBanana) firstBanana = entry;
      }
    }

    // 🖊 Записуємо у верхній блок (банани)
    if (totals.banana.qty > 0) {
      const e = firstBanana || entries[0];
      mainSheet.getCell('J8').value = e['Data wysyłki'] || '';
      mainSheet.getCell('C8').value = client || '';
      mainSheet.getCell('J25').value = `${totals.banana.qty} (${totals.banana.pal})`;
      mainSheet.getCell('J29').value = e['Kierowca'] || '';
      mainSheet.getCell('E10').value = e['Godzina'] || '';
    }

    // 🖊 Записуємо у нижній блок (BIO) на тому ж аркуші
    if (totals.bio.qty > 0) {
      const e = firstBio || entries[0];
      mainSheet.getCell('J60').value = e['Data wysyłki'] || '';
      mainSheet.getCell('C60').value = `${client} (BIO)`;
      mainSheet.getCell('J69').value = `${totals.bio.qty} (${totals.bio.pal})`;
      mainSheet.getCell('K63').value = e['Kierowca'] || '';
      mainSheet.getCell('E61').value = e['Godzina'] || '';
    }

    // 🔍 Лог для контролю класифікації
    console.log(`➡️  ${client} [${truck}]  banana=${totals.banana.qty} / bio=${totals.bio.qty}`);

    // 📂 Збереження файлу
    const safeClientName = client.replace(/[\\/:*?"<>|]/g, '_');
    const safeTruck = truck.replace(/[\\/:*?"<>|]/g, '_');
    const clientBaseDir = path.join(outputDir, safeClientName);
    if (!fs.existsSync(clientBaseDir)) fs.mkdirSync(clientBaseDir, { recursive: true });

    const fileName = `Quality report ${safeClientName}_${safeTruck}.xlsx`;
    const outputPath = path.join(clientBaseDir, fileName);

    await workbook.xlsx.writeFile(outputPath);
    console.log(`📄 Створено файл: ${outputPath}`);
  }

  console.log('✅ Усі звіти згенеровано успішно!');
}

fillTemplate().catch(err => {
  console.error('❌ Помилка:', err);
  process.exit(1);
});
