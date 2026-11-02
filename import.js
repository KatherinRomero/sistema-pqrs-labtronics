// Importa las respuestas existentes: npm run importar -- ruta/REPORTE_DE_PQRS_(Respuestas).xlsx
const ExcelJS = require('exceljs'), crypto = require('crypto'), { db } = require('./db');
(async () => {
  const w = new ExcelJS.Workbook(); await w.xlsx.readFile(process.argv[2]);
  const ws = w.getWorksheet('Respuestas de formulario 1'), ins = db.prepare('insert or ignore into p(id,num,codigo,ts,pos,fin,d) values(?,?,?,?,0,?,?)');
  let n = 0;
  ws.eachRow((row, i) => {
    if (i < 2 || !(row.getCell(1).value instanceof Date)) return;
    const d = {};
    for (let c = 2; c <= 77; c++) {
      let v = row.getCell(c).value;
      if (v == null || (typeof v === 'object' && !(v instanceof Date))) continue;
      v = v instanceof Date ? v.toISOString().slice(0, 10) : String(v).trim();
      if (v) d[c] = v;
    }
    ins.run(crypto.randomUUID(), i, 'PQRS-' + String(i).padStart(4, '0'), row.getCell(1).value.toISOString(), d[73] || d[75] ? 1 : 0, JSON.stringify(d)); n++;
  });
  console.log(n + ' PQRS importadas');
})();
