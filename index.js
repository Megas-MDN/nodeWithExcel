const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs/promises');
const assets = require('./assets/__assets.js');
// create a xlsx file
const workbook = xlsx.utils.book_new();

const genObj = (ticket) => ({
  ativo: ticket,
  'Quantidade de Compra': `=RTD("tryd.rtdserver";;"COT";"${ticket}";"QtdCpa")`,
  Compra: `=RTD("tryd.rtdserver";;"COT";"${ticket}";"PrcCpa")`,
  Ultimo: `=RTD("tryd.rtdserver";;"COT";"${ticket}";"Ult")`,
  Venda: `=RTD("tryd.rtdserver";;"COT";"${ticket}";"PrcVda")`,
  'Quantidade de Venda': `=RTD("tryd.rtdserver";;"COT";"${ticket}";"QtdVda")`,
  Negocios: `=RTD("tryd.rtdserver";;"COT";"${ticket}";"QtdNeg")`,
});

fs.readFile(path.resolve(assets, 'dataTickets.json'), 'utf8').then((data) => {
  const tickets = JSON.parse(data);
  const ws = xlsx.utils.json_to_sheet(tickets);

  xlsx.utils.book_append_sheet(workbook, ws, 'Sheet1');

  xlsx.writeFile(workbook, path.resolve(assets, 'data.xlsx'));
  console.log('Done ---');
});
