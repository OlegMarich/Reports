const {exec} = require('child_process');
const path = require('path');

const dateArg = process.argv[2];
if (!dateArg) {
  console.error('❌ Не передано дату як аргумент');
  process.exit(1);
}

const generateScript = path.join(__dirname, 'generate-reports.js');
const loadingScript = path.join(__dirname, 'fill-template-loading.js');
const templateScript = path.join(__dirname, 'fill-template-client.js');
const cleanScript = path.join(__dirname, 'fill-template-clean.js');
const shippingScript = path.join(__dirname, 'fill-template-shipping.js');

console.log('🚀 Generating report...');

exec(`node "${generateScript}" ${dateArg}`, (err, stdout, stderr) => {
  if (err) {
    console.error('❌ Error during generate-reports:', stderr || err.message);
    process.exit(1);
  }
  console.log(stdout);

  console.log('📦 Filling loading template...');
  exec(`node "${loadingScript}" ${dateArg}`, (err2, stdout2, stderr2) => {
    if (err2) {
      console.error('❌ Error during loading-template:', stderr2 || err2.message);
      process.exit(1);
    }
    console.log(stdout2);

    console.log('📦 Filling client templates...');
    exec(`node "${templateScript}" ${dateArg}`, (err3, stdout3, stderr3) => {
      if (err3) {
        console.error('❌ Error during client-template:', stderr3 || err3.message);
        process.exit(1);
      }
      console.log(stdout3);

      console.log('📦 Filling clean template...');
      exec(`node "${cleanScript}" ${dateArg}`, (err4, stdout4, stderr4) => {
        if (err4) {
          console.error('❌ Error during clean-template:', stderr4 || err4.message);
          process.exit(1);
        }
        console.log(stdout4);

        console.log('@@@DONE:' + dateArg);
      });
    });
  });
});
