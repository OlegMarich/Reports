const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const selectBtn = document.getElementById('selectBtn');
const generateBtn = document.getElementById('generateBtn');
const dateInput = document.getElementById('reportDate'); // для звіту
const weekInput = document.getElementById('weekInput');   // для плану

const modal = document.getElementById('modalMessage');
const modalText = document.getElementById('modalMessageText');
const modalOkBtn = document.getElementById('modalOkBtn');

const modalWeekPrompt = document.getElementById('modalWeekPrompt'); // тільки для plan

let selectedFiles = [];
let generatedPath = null;

// 🟡 Визначаємо активний режим зі <body data-mode="...">
function getMode() {
  return document.body.dataset.mode?.toLowerCase();
}

// 📁 Вибір файлів
selectBtn?.addEventListener('click', () => {
  fileInput?.click();
});

fileInput?.addEventListener('change', (event) => {
  selectedFiles = Array.from(event.target.files);
  if (selectedFiles.length > 0) {
    generateBtn.style.display = 'inline-block';
    dropZone.innerHTML = `Selected: <strong>${selectedFiles.map((f) => f.name).join(', ')}</strong>`;
  }
});

// ▶️ Обробка кнопки Generate
generateBtn?.addEventListener('click', () => {
  const mode = getMode();

  if (mode === 'plan') {
    // Показати модалку для тижня
    weekInput.value = '';
    modalWeekPrompt.style.display = 'flex';
  } else if (mode === 'report') {
    // Запустити щоденний звіт
    handleDailyReport();
  } else {
    alert('❌ Unknown mode. Set <body data-mode="plan"> або "report".');
  }
});

// ❌ Закрити модалку тижня
document.getElementById('cancelWeekBtn')?.addEventListener('click', () => {
  modalWeekPrompt.style.display = 'none';
});

// ✅ Підтвердження тижня — запуск генерації плану
document.getElementById('confirmWeekBtn')?.addEventListener('click', async () => {
  const week = weekInput.value.trim();
  if (!week || isNaN(week) || Number(week) < 1 || Number(week) > 53) {
    return alert('⚠️ Please enter a valid week number (1–53).');
  }

  modalWeekPrompt.style.display = 'none';
  generateBtn.disabled = true;
  generateBtn.textContent = 'Generating...';

  try {
    const response = await fetch(`/api/generate-plan?week=${week}`);
    const text = await response.text();

    let result;
    try {
      result = JSON.parse(text);
    } catch (err) {
      console.error('❌ Not a JSON:', text);
      showModalMessage('❌ Server did not return valid JSON. Check logs.');
      return;
    }

    if (result.message?.includes('✅')) {
      generatedPath = `${week}_Week`;
      showModalMessage(`✅ Plan for <strong>week ${week}</strong> generated.<br>Check <code>/output/${generatedPath}</code>`);
    } else {
      showModalMessage(result.message || '❌ Failed to generate plan.');
    }
  } catch (err) {
    console.error(err);
    showModalMessage('❌ Something went wrong.');
  } finally {
    generateBtn.disabled = false;
    generateBtn.textContent = 'Generate Report';
  }
});

// 📅 Генерація щоденного звіту
async function handleDailyReport() {
  const date = dateInput?.value;

  if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    return showModalMessage('⚠️ Please select a valid date (YYYY-MM-DD).');
  }

  if (selectedFiles.length === 0) {
    return showModalMessage('⚠️ Please select both Excel files.');
  }

  const formData = new FormData();
  selectedFiles.forEach((file) => formData.append('files', file));

  try {
    generateBtn.disabled = true;
    generateBtn.textContent = 'Generating...';

    const response = await fetch(`/upload?date=${date}`, {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

    const result = await response.json();

    if (result.success) {
      generatedPath = result.date;
      showModalMessage(`✅ Report for <strong>${result.date}</strong> generated.<br>See <code>/output/${result.date}</code>`);
    } else {
      showModalMessage('❌ Failed to generate report.');
    }
  } catch (err) {
    console.error(err);
    showModalMessage('❌ Something went wrong.');
  } finally {
    generateBtn.disabled = false;
    generateBtn.textContent = 'Generate Report';
  }
}

// 📦 Показ повідомлення
function showModalMessage(message) {
  modalText.innerHTML = message;
  modal.style.display = 'flex';
}

// 🪟 Закриття модалки → відкриття папки
modalOkBtn?.addEventListener('click', () => {
  modal.style.display = 'none';
  if (generatedPath) {
    window.open(`/output/${generatedPath}`, '_blank');
  }
});
