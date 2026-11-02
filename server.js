const express = require('express'), nodemailer = require('nodemailer'), ExcelJS = require('exceljs'), puppeteer = require('puppeteer'),
  fs = require('fs'), crypto = require('crypto'), path = require('path'), { db, D } = require('./db');
const S = require('./schema.json'), H = require('./headers.json'), E = process.env, Q = S.secs.flatMap(s => s.qs);
const BASE = (E.BASE_URL || 'http://localhost:3000').replace(/\/$/, '');
const mail = nodemailer.createTransport({ host: E.SMTP_HOST, port: +(E.SMTP_PORT || 587), secure: E.SMTP_PORT == 465, auth: { user: E.SMTP_USER, pass: E.SMTP_PASS } });
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const get = r => { const x = db.prepare('select * from p where id=? or codigo=?').get(r, String(r).toUpperCase()); if (!x) throw Error('PQRS no encontrada'); x.d = JSON.parse(x.d); return x; };
const auth = (r, c) => { if (String(c || '').trim() !== String(E.CODIGO_EQUIPO || 'Labtronics2026').trim()) throw Error('Código incorrecto. Solo el equipo de Labtronics SAS puede continuar'); return get(r); };
const put = x => db.prepare('update p set d=?,pos=?,fin=? where id=?').run(JSON.stringify(x.d), x.pos, x.fin, x.id);
const link = id => { const x = db.prepare('select codigo from p where id=? or codigo=?').get(id, String(id).toUpperCase()); return `${BASE}/${x ? x.codigo : ''}`; };
const LOG = 'liderlogistica@labtronics.net,almacen@labtronics.net';
const RUTAS = {
  'Desempeño del producto (Reactivo/Insumo)': ['especialistadeproducto1@labtronics.net'],
  'Desempeño del equipo (Analizadores/Repuestos)': ['coordinadoringenieria@labtronics.net', 'ingenieria@labtronics.net'],
  'Almacén (Pérdida de la cadena de frío/Faltantes en el pedido)': [LOG],
  'Logística (Estado del producto/Tiempos de entrega)': [LOG],
  'Errores de facturación (faltantes /sobrantes)': ['pedidos@labtronics.net'],
  'Otros': ['infolab@labtronics.net']
};
const SKIP = /exclusiva para el equipo|usted hace parte|se han realizado todas/i;

const ACT = S.secs.find(s => s.loop).qs, DESC = [17, 24, 26, 27];
// Columna/clave de la pregunta j de la acción k (las 5 primeras usan las columnas originales; el resto, claves extra)
const ak = (k, j) => (k <= 4 || (k == 5 && j < 4)) ? 42 + 5 * (k - 1) + j : `x${k}_${j}`;
const acts = d => { let n = 0; for (let k = 1; k <= 99; k++) if (ACT.some((q, j) => d[ak(k, j)])) n = k; return n; };
const secsOf = d => S.secs.flatMap(s => s.loop ? Array.from({ length: acts(d) }, (_, i) => ({ title: `SOLUCIÓN · ACCIÓN ${i + 1}`,
  qs: ACT.map((q, j) => ({ ...q, c: ak(i + 1, j), title: `${q.title} (acción ${i + 1})` })) })) : [s]);
function cuerpo(x, nueva) {
  let t = nueva ? 'Buen dia,\n\nMe permito informar que se genero una nueva PQRS.\n\n'
    : 'Buen dia,\n\nMe permito informarle que la PQRS ha sido finalizada.\nA continuacion encontrara las respuestas registradas y, al final, el PDF de la PQRS.\n\n';
  t += `Con codigo: ${x.codigo}\n`;
  if (nueva) t += `\nContinua con la trazabilidad de la PQRS por medio de este enlace:\n${link(x.id)}\n\n`;
  secsOf(x.d).flatMap(s => s.qs).forEach(q => {
    const v = x.d[q.c]; if (q.t == 'k' || !v) return;
    t += `${q.title}:\n` + (q.t == 'f' ? 'Archivos adjuntos:\n' + v.split(', ').map(u => 'Adjunto: ' + u).join('\n') : v) + '\n\n';
  });
  return t;
}
function pdfHtml(x) {
  const body = secsOf(x.d).map(s => {
    const qs = s.qs.filter(q => q.t != 'k' && !SKIP.test(q.title) && x.d[q.c]); if (!qs.length) return '';
    return `<div class=c><h2>${esc(s.title || 'Gastos')}</h2><div class=g>` + qs.map(q => `<div class=e><div class=k>${esc(q.title)}</div>` +
      (q.t == 'f' ? x.d[q.c].split(', ').map((u, i) => `<a href="${esc(u)}">Archivo ${i + 1}</a>`).join('<br>') : `<p>${esc(x.d[q.c]).replace(/\n/g, '<br>')}</p>`) + '</div>').join('') + '</div></div>';
  }).join('');
  return `<!doctype html><meta charset=utf-8><style>@page{size:A4;margin:6mm}body{font:9px Arial;color:#1b1f24}h1{font-size:17px;color:#1f3b5c;margin:0}
.m{width:100%;border-collapse:collapse;margin:8px 0}.m td{border:1px solid #d8e2ee;padding:3px 5px}.m .h{background:#f2f7ff;font-weight:bold;color:#1f3b5c;width:80px}
.c{border:1px solid #d8e2ee;border-radius:6px;margin:0 0 8px;overflow:hidden}h2{background:#d8e8ff;color:#0c3b76;font-size:10px;margin:0;padding:7px 10px;text-transform:uppercase}
.g{columns:2}.e{break-inside:avoid;padding:4px 6px}.k{background:#f8fbff;font-weight:bold;color:#244865;padding:3px}p{margin:2px 0}a{color:#1155cc}</style>
<h1>Cierre de PQRS</h1><div>Resumen completo de respuestas registradas · Código CO-FR-010 · Versión 4 · Fecha 2024-11-08</div>
<table class=m><tr><td class=h>Codigo</td><td>${esc(x.codigo)}</td><td class=h>Cliente</td><td>${esc(x.d[4] || 'Sin nombre')}</td></tr>
<tr><td class=h>Tipo de caso</td><td>${esc(x.d[10] || '')}</td><td class=h>ID respuesta</td><td>${esc(x.id)}</td></tr></table>${body}`;
}
async function pdf(x) {
  const b = await puppeteer.launch({ args: ['--no-sandbox'], executablePath: E.PUPPETEER_EXECUTABLE_PATH || undefined });
  try {
    const pg = await b.newPage(); await pg.setContent(pdfHtml(x));
    const buf = Buffer.from(await pg.pdf({ format: 'A4', printBackground: true }));
    fs.writeFileSync(`${D}/pdf/${x.id}.pdf`, buf); return buf;
  } finally { await b.close(); }
}
async function enviar(x, nueva) {
  const area = x.d[10] || '', r = RUTAS[area] || [''];
  const o = { from: E.MAIL_FROM || E.SMTP_USER, to: [r[0], 'asistentecomercial@labtronics.net'].filter(Boolean).join(','), cc: r[1] || '',
    subject: nueva ? `Codigo PQRS ${x.codigo} de tipo: ${area || 'Sin tipo'}` : `Se finalizo la PQRS ${x.codigo}${x.d[4] ? ' - ' + x.d[4] : ''}`, text: cuerpo(x, nueva) };
  if (!nueva) { o.attachments = [{ filename: `Cierre_${x.codigo}.pdf`, content: await pdf(x) }]; o.text += `Al final encontrara el PDF de la PQRS en el siguiente enlace:\n${BASE}/pdf/${x.id}.pdf\n\n`; }
  await mail.sendMail(o);
}
async function excel() {
  const w = new ExcelJS.Workbook(), a = w.addWorksheet('Respuestas de formulario 1'), c = w.addWorksheet('CONSOLIDADO DE QUEJAS, RECLAMOS');
  const rows = db.prepare('select * from p order by num').all().map(x => ({ x, d: JSON.parse(x.d) }));
  const ex = [...new Set(rows.flatMap(r => Object.keys(r.d).filter(k => /^x\d+_\d$/.test(k))))].sort((p, q) => p.localeCompare(q, undefined, { numeric: true }));
  a.addRow([...H, ...ex.map(k => { const m = /x(\d+)_(\d)/.exec(k); return `Acción ${m[1]}: ${ACT[m[2]].title}`; })]);
  rows.forEach(({ x, d }) => {
    const r = new Array(H.length).fill(''); r[0] = new Date(x.ts);
    Object.keys(d).forEach(k => { if (/^\d+$/.test(k)) r[k - 1] = d[k]; }); r[H.length - 1] = x.codigo;
    a.addRow([...r, ...ex.map(k => d[k] || '')]);
  });
  c.addRow(['CONSOLIDADO DE QUEJAS, RECLAMOS Y SUGERENCIAS']);
  c.addRow(['No', 'FECHA', 'TIPO', 'NOMBRE DEL CLIENTE', 'PROCESO RESPONSABLE', 'DESCRIPCIÓN DE LA PQRS', 'PRIMER CONTACTO CON EL CLIENTE', 'INVESTIGACIÓN / CAUSAS PROBABLES', 'PROCEDENTE', 'SOLUCIÓN DE LA PQRS (Acciones ejecutadas)', 'FECHA DE SOLUCIÓN', 'ESTADO', 'CALIFICACIÓN DE LA GESTIÓN', 'REQUIERE ACCIÓN', 'TIEMPO DE SOLUCIÓN (días)']);
  filas().forEach(f => c.addRow([f.no, new Date(f.fecha), f.tipo, f.cliente, f.proceso, f.descripcion, f.contacto, f.investigacion, f.procedente, f.solucion, f.fechaSol, f.estado, f.cal, f.acc, f.dias]));
  c.columns.forEach(col => { col.width = 28; col.alignment = { wrapText: true, vertical: 'top' }; });
  const f = path.resolve(`${D}/PQRS_Respuestas.xlsx`); await w.xlsx.writeFile(f); return f;
}
// Filas del consolidado (la misma fuente alimenta la pagina online y el Excel)
function filas() {
  return db.prepare('select * from p order by num').all().map(x => {
    const d = JSON.parse(x.d), n = acts(d), g = (k, j) => d[ak(k, j)] || '';
    const solucion = Array.from({ length: n }, (_, i) => `${i + 1}) ${g(i + 1, 1)} · ${g(i + 1, 0)}: ${g(i + 1, 2)}${g(i + 1, 3) ? ' → ' + g(i + 1, 3) : ''}`).join('\n');
    return { id: x.id, no: x.codigo, fecha: x.ts, tipo: d[2] || d[10] || '', cliente: [4, 5, 6, 7, 8].map(c => d[c]).filter(Boolean).join('\n'),
      proceso: d[34] || '', descripcion: DESC.map(c => d[c]).find(Boolean) || '', contacto: [d[31], d[32]].filter(Boolean).join(' · '),
      investigacion: d[38] || d[40] || '', procedente: d[35] || '', solucion, fechaSol: d[75] || '', estado: x.fin ? 'Cerrado' : 'Abierto',
      cal: d.cal || '', acc: d.acc || '', dias: d[75] ? Math.max(0, Math.round((new Date(d[75]) - new Date(x.ts)) / 864e5)) : '',
      pdf: x.fin ? `${BASE}/pdf/${x.id}.pdf` : '', link: link(x.id) };
  });
}
const API = {
  api_crear: async d => {
    const n = Math.max(db.prepare('select max(num) m from p').get().m || 0, +(E.CONTADOR_INICIAL || 3)) + 1, dd = {};
    Q.forEach(q => { if (q.p == 1 && d[q.c]) dd[q.c] = String(d[q.c]); });
    const x = { id: crypto.randomUUID(), num: n, codigo: 'PQRS-' + String(n).padStart(4, '0'), ts: new Date().toISOString(), pos: 0, fin: 0, d: dd };
    db.prepare('insert into p(id,num,codigo,ts,pos,fin,d) values(?,?,?,?,0,0,?)').run(x.id, n, x.codigo, x.ts, JSON.stringify(dd));
    let errorCorreo = ''; try { await enviar(x, true); } catch (e) { errorCorreo = String(e); console.error(e); }
    excel().catch(console.error);
    return { r: x.id, codigo: x.codigo, url: link(x.id), errorCorreo };
  },
  api_entrar: async (r, c) => { const x = auth(r, c); x.d[Q.find(q => q.t == 'k').c] = c; put(x); return { v: x.d, url: link(x.id) }; },
  api_guardar: async (r, c, col, v) => {
    const x = auth(r, c);
    if (!(/^x\d{1,2}_[0-4]$/.test(col) || Q.some(q => q.c == col && q.p == 2 && q.t != 'k'))) throw Error('Columna no permitida');
    v ? x.d[col] = String(v) : delete x.d[col]; put(x); return 1;
  },
  api_subir: async (n, m, b) => {
    if (b.length > 14e6) throw Error('Archivo mayor a 10 MB');
    const id = crypto.randomUUID() + '_' + String(n).replace(/[^\w.-]/g, '_');
    fs.writeFileSync(`${D}/files/${id}`, Buffer.from(b, 'base64')); return `${BASE}/f/${id}`;
  },
  api_finalizar: async (r, c) => {
    const x = auth(r, c);
    if (!secsOf(x.d).flatMap(s => s.qs).some(q => /firma/i.test(q.title) && x.d[q.c])) return { ok: false, msg: 'Adjunte el archivo de su firma para finalizar la PQRS.' };
    await enviar(x, false); x.fin = 1; put(x); excel().catch(console.error); return { ok: true, codigo: x.codigo };
  }
};
const app = express();
app.use(express.json({ limit: '25mb' }));
const seguro = { setHeaders: r => { r.setHeader('X-Content-Type-Options', 'nosniff'); r.setHeader('Content-Security-Policy', 'sandbox'); } };
app.use('/f', express.static(`${D}/files`, seguro)); app.use('/pdf', express.static(`${D}/pdf`, seguro));
app.get('/api/schema', (q, r) => r.json(S));
app.post('/api/:fn', async (q, r) => {
  try { const f = API[q.params.fn]; if (!f) throw Error('No existe'); r.json(await f(...q.body)); }
  catch (e) { r.status(400).json({ error: e.message }); }
});
// Panel online del consolidado (usuario/clave de .env)
const adm = (q, r, n) => {
  const [u, p] = Buffer.from((q.headers.authorization || '').slice(6), 'base64').toString().split(':');
  if (!E.ADMIN_PASS || u !== E.ADMIN_USER || p !== E.ADMIN_PASS) return r.set('WWW-Authenticate', 'Basic').status(401).end(); n();
};
app.get('/admin', adm, (q, r) => r.sendFile(path.resolve('admin.html')));
app.get('/admin/data', adm, (q, r) => r.json(filas()));
app.post('/admin/set', adm, (q, r) => {
  const { id, k, v } = q.body; if (!['cal', 'acc'].includes(k)) return r.status(400).end();
  const x = get(id); v ? x.d[k] = String(v) : delete x.d[k]; put(x); excel().catch(console.error); r.json(1);
});
app.get('/admin/excel', adm, async (q, r) => r.download(await excel(), 'REPORTE DE PQRS (Respuestas).xlsx'));
app.get(/^\/pqrs-\d+$/i, (q, r) => r.sendFile(path.resolve('public/index.html')));  // enlace de seguimiento: /PQRS-0004
app.use(express.static('public'));
app.listen(E.PORT || 3000, () => console.log('PQRS en puerto ' + (E.PORT || 3000)));
