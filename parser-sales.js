const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const inquirer = require('inquirer');

// 📥 Зчитує файл SALES + повертає вибрану вкладку як масив рядків
async function readSalesPlan() {
  const inputDir = path.join(__dirname, 'input');

  const salesFile = fs.readdirSync(inputDir).find((f) => f.toLowerCase().includes('sales'));

  if (!salesFile) {
    console.error('❌ Не знайдено файл sales у папці /input');
    process.exit(1);
  }

  const salesPath = path.join(inputDir, salesFile);
  const workbook = xlsx.readFile(salesPath);

  const { weekName } = await inquirer.prompt({
    type: 'list',
    name: 'weekName',
    message: '🗓 Оберіть вкладку з тижнем:',
    choices: workbook.SheetNames,
  });

  const sheet = workbook.Sheets[weekName];
  if (!sheet) {
    console.error('❌ Вкладка не знайдена!');
    process.exit(1);
  }

  const sheetJson = xlsx.utils.sheet_to_json(sheet, {
    defval: '',
    header: 1,
  });

  return { sheetJson, weekName, salesFile, sheet };
}

function extractDatesFromHeader(sheet) {
  const XLSX = require('xlsx');

  // Отримаємо всі адреси комірок (наприклад, A1, B3, C5)
  const cellAddresses = Object.keys(sheet).filter((addr) => addr[0] !== '!');

  // Фільтруємо, щоб взяти лише ті, що знаходяться в рядку 3
  const row3Cells = cellAddresses.filter((addr) => {
    const match = addr.match(/^([A-Z]+)(\d+)$/);
    return match && match[2] === '3';
  });

  // Сортуємо по колонках: A, B, C...
  row3Cells.sort((a, b) => {
    const colToNum = (col) =>
      col.split('').reduce((res, ch) => res * 26 + (ch.charCodeAt(0) - 64), 0);
    const colA = a.match(/^([A-Z]+)/)[1];
    const colB = b.match(/^([A-Z]+)/)[1];
    return colToNum(colA) - colToNum(colB);
  });

  const dates = row3Cells
    .map((addr) => {
      const value = sheet[addr]?.v;
      if (typeof value === 'string') {
        const match = value.match(/^(\d{1,2})[-\/](\d{1,2})$/);
        if (match) {
          const day = match[1].padStart(2, '0');
          const month = match[2].padStart(2, '0');
          const now = new Date();
          const year = now.getFullYear();
          return `${year}-${month}-${day}`;
        }
      }

      if (typeof value === 'number') {
        const dateObj = XLSX.SSF.parse_date_code(value);
        if (dateObj) {
          const year = dateObj.y;
          const month = String(dateObj.m).padStart(2, '0');
          const day = String(dateObj.d).padStart(2, '0');
          return `${year}-${month}-${day}`;
        }
      }

      return null;
    })
    .filter(Boolean);

  return dates;
}

const clientsData = {
  'penny karcag': { product: 'BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: 'YFF HU' },
  'penny karcag bio bananas': { product: 'BIO BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: 'YFF RO' },
  'penny veszprem': { product: 'BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: 'Citronex' },
  'penny veszprem bio bananas': { product: 'BIO BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: 'forwarder' },
  'penny alsonemedi': { product: 'BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: '' },
  'penny alsonemedi bio bananas': { product: 'BIO BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: '' },
  'aldi biatorbagy banana': { product: 'BANANA', 'box/pal': 28, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: '' },
  'aldi biatorbagy ananas': { product: 'ANANAS', 'box/pal': 40, 'weight/box': 12.95, 'pal type': 'EUR', forwarder: '' },
  'spar ullo': { product: 'BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: '' },
  'spar bicske': { product: 'BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: '' },
  'ta-moro kft.': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': '', forwarder: '' },
  'billa senec': { product: 'BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: '' },
  'billa říčany': { product: 'BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: '' },
  'billa petrovany': { product: 'BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: '' },
  'yff kft. - hu inb. bananas': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'yff kft. - hu inb. ananas': { product: 'ANANAS', 'box/pal': 80, 'weight/box': 12.95, 'pal type': 'industrial', forwarder: '' },
  'yff srl - remetea': { product: 'BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: '' },
  'yff srl - kaufland turda': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: '' },
  'studenac dugopolje': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'studenac zagreb': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'citronex zgorzelec': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'yff kft. - spar hrvatska klinča sela': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: '' },
  'yff kft - spar ljubljana': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'kaufland gliwice': { product: 'BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'fruit expert bytomska': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'yff kft. - aldi lukovica #2': { product: 'BANANA', 'box/pal': 2, 'weight/box': 20.29, 'pal type': 'PHP mini', forwarder: '' },
  'yff kft. - aldi lukovica #4': { product: 'BANANA', 'box/pal': 4, 'weight/box': 20.29, 'pal type': 'PHP mini', forwarder: '' },
  'yff kft. - aldi lukovica #8': { product: 'BANANA', 'box/pal': 8, 'weight/box': 20.29, 'pal type': 'PHP mini', forwarder: '' },
  'm & w frischgemuse wien': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'lunys s.r.o. - bratislava': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'lunys s.r.o. - poprad': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'jmp biedronka,s. r. o. banana': { product: 'BANANA', 'box/pal': 28, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: '' },
  'jmp biedronka,s. r. o. tomatoes': { product: 'BANANA', 'box/pal': 72, 'weight/box': 7, 'pal type': 'EUR', forwarder: '' },
  'partner log kft cba': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'yff kft. - aldi lukovica tomato': { product: 'TOMATO', 'box/pal': 56, 'weight/box': 7, 'pal type': 'EUR', forwarder: '' },
  'yff kft. - hu inb. tomatoes': { product: 'TOMATO', 'box/pal': 72, 'weight/box': 7, 'pal type': 'industrial', forwarder: '' },
  'frutura obst & gemuse hartl, austria': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'lidl hu - szekesfehervar': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'lidl hu - hejokurt': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'lidl hu - ecser': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'lidl hu - szigetszentmiklos': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'yff kft - metro': { product: 'BANANA', 'box/pal': 32, 'weight/box': 19.79, 'pal type': 'EUR', forwarder: '' },
  'veradel kft.': { product: 'BANANA', 'box/pal': 1, 'weight/box': 19.79, 'pal type': '', forwarder: '' },
  'greenyard fresh austria gmbh': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'ppo siechnice sp.zo.o': { product: 'TOMATO', 'box/pal': 56, 'weight/box': 7, 'pal type': 'EUR', forwarder: '' },
  'ivanyi zoldsegkert': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': 'industrial', forwarder: '' },
  'kaufland sk ilava': { product: 'BANANA', 'box/pal': 48, 'weight/box': 19.79, 'pal type': '', forwarder: '' }
};


function parseSalesByCustomer(sheetJson, dates) {
  const result = [];
  let currentCustomer = null;
  let currentColorCode = null;
  let currentCityOrLocation = null;

  const startDateCol = 2; // Колонка з датами починається з індексу 2 (тобто "D")

  for (let i = 11; i < sheetJson.length; i++) {
    const row = sheetJson[i];

    // Пропускаємо порожні рядки
    if (!row || row.every((cell) => cell === '' || cell === undefined)) {
      continue;
    }

    // Спроба знайти назву клієнта
    const candidate = row.slice(0, 5).find(
      (cell) => typeof cell === 'string' && cell.trim().length > 2 && !cell.trim().match(/^\d+$/),
    );

    const onlyText = row.every((cell) => typeof cell === 'string' || cell === '');

    if (candidate && onlyText) {
      currentCustomer = candidate.trim();

      // Витягуємо колір з дужок
      const match = currentCustomer.match(/\[(.*?)\]/);
      currentColorCode = match ? match[1].trim() : null;

      // Скидаємо поточне місто/локацію при новому клієнті
      currentCityOrLocation = null;

      continue;
    }

    // Перевірка, чи в рядку є числові замовлення
    const quantities = row.slice(startDateCol, startDateCol + dates.length);
    const hasNumbers = quantities.some((v) => typeof v === 'number' && !isNaN(v));

    if (currentCustomer && hasNumbers) {
      // Пропускаємо записи з total
      if (
        currentCustomer.trim().toLowerCase() === 'total' ||
        (row[0] && row[0].toString().trim().toLowerCase().includes('total'))
      ) {
        continue;
      }

      const firstCell = (row[0] || '').toString().trim();

      let lineName;
      if (firstCell && !/\d/.test(firstCell)) {
        if (/bio/i.test(firstCell)) {
          lineName = currentCityOrLocation ? `${currentCityOrLocation} ${firstCell}` : firstCell;
        } else {
          currentCityOrLocation = firstCell;
          lineName = currentCityOrLocation;
        }
      } else {
        lineName = 'Unknown';
      }

      // Ще раз фільтруємо по lineName
      if (lineName.trim().toLowerCase().includes('total')) {
        continue;
      }

      const data = dates.map((date, idx) => ({
        date,
        qty: quantities[idx] || 0,
      }));

      result.push({
        customer: currentCustomer,
        colorCode: currentColorCode,
        line: lineName,
        data,
      });
    }
  }

  return result;
}

async function main() {
  const { sheetJson, weekName, salesFile, sheet } = await readSalesPlan();
  const dates = extractDatesFromHeader(sheet);
  const parsed = parseSalesByCustomer(sheetJson, dates);

  const outputDir = path.join(__dirname, 'output', weekName.replace(/\s+/g, '_'));
  fs.mkdirSync(outputDir, { recursive: true });

  const outputPath = path.join(outputDir, 'sales.json');
  fs.writeFileSync(outputPath, JSON.stringify(parsed, null, 2), 'utf8');

  console.log(`✅ Дані збережено у: ${outputPath}`);
  console.log('📋 Вкладка:', weekName);
  console.log('🔢 Рядків зчитано:', sheetJson.length);
  console.log('📅 Знайдені дати:', dates);
  console.log('🔍 Знайдено клієнтських рядків:', parsed.length);
  console.dir(parsed.slice(0, 5), { depth: null });
}

main();
