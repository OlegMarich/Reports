const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { exec } = require('child_process');

const app = express();
const PORT = 3000;

const inputDir = path.join(__dirname, 'input');
const outputDir = path.join(__dirname, 'output');
const publicDir = path.join(__dirname, 'public');

// Статичні файли
app.use(express.static(publicDir));
app.use('/output', express.static(outputDir));

// Налаштування збереження файлів для multer
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, inputDir),
  filename: (req, file, cb) => cb(null, file.originalname),
});
const upload = multer({ storage });

// Створюємо output, якщо не існує
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

// =====================
// Маршрут для генерації плану по GET /api/generate-plan?week=1
app.get('/api/generate-plan', (req, res) => {
  const week = req.query.week;
  if (!week) return res.status(400).json({ message: '❌ Week number missing' });

  // Формуємо назву вкладки
  const sheetName = `${week} WEEK`;

  exec(`node parser-sales.js "${sheetName}"`, (err, stdout, stderr) => {
    if (err) {
      console.error('❌ Parser error:', err.message);
      console.error('stderr:', stderr);
      return res.status(500).json({ message: stderr || '❌ Parser script failed' });
    }

    exec(`node fill-template-plan.js ${week}`, (err2, stdout2, stderr2) => {
      if (err2) {
        console.error('❌ Plan error:', err2.message);
        console.error('stderr:', stderr2);
        return res.status(500).json({ message: stderr2 || '❌ Plan script failed' });
      }

      console.log('✅ Plan script output:', stdout2);
      res.json({ message: stdout2.includes('✅') ? stdout2 : '✅ Plan generated.' });
    });
  });
});

// =====================
// Маршрут для завантаження файлів і запуску генерації звітів по POST /upload?date=YYYY-MM-DD
app.post('/upload', upload.array('files', 2), (req, res) => {
  const userDate = req.query.date;

  if (!userDate || !/^\d{4}-\d{2}-\d{2}$/.test(userDate)) {
    return res.status(400).json({ success: false, message: 'Invalid or missing date parameter' });
  }

  const cmd = `node run-all.js ${userDate}`;
  exec(cmd, (err, stdout, stderr) => {
    if (err) {
      console.error('❌ Error during script run:', stderr);
      return res.status(500).json({ success: false, message: 'Script execution error' });
    }

    console.log(stdout);

    const match = stdout.match(/@@@DONE:(\d{4}-\d{2}-\d{2})/);
    const resultDate = match ? match[1] : null;

    if (!resultDate) {
      return res.status(500).json({ success: false, message: 'No completion confirmation found' });
    }

    // Відкриваємо папку output у файловому провіднику Windows
    const folderPath = path.join(outputDir, resultDate);
    exec(`start "" "${folderPath}"`, (openErr) => {
      if (openErr) {
        console.error('❌ Error opening folder:', openErr);
      }
    });

    res.json({
      success: true,
      message: 'Report generated successfully',
      date: resultDate,
    });
  });
});

// =====================
// Запуск сервера
app.listen(PORT, () => {
  console.log(`🚀 Server running at http://localhost:${PORT}`);
});
