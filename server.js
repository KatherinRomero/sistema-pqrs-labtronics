const express = require('express');
const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');
const crypto = require('crypto');
const path = require('path');
const { supabase } = require('./db');

const S = require('./schema.json');
const H = require('./headers.json');
const E = process.env;
const Q = S.secs.flatMap(s => s.qs);
const BASE = (E.BASE_URL || 'http://localhost:3000').replace(/\/$/, '');
const BUCKET = E.SUPABASE_BUCKET || 'pqrs-documentos';
const mail = nodemailer.createTransport({
  host: E.SMTP_HOST,
  port: +(E.SMTP_PORT || 587),
  secure: E.SMTP_PORT == 465,
  auth: E.SMTP_USER && E.SMTP_PASS ? { user: E.SMTP_USER, pass: E.SMTP_PASS } : undefined
});

const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
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
const ACT = S.secs.find(s => s.loop).qs;
const DESC = [17, 24, 26, 27];
const ak = (k, j) => (k <= 4 || (k === 5 && j < 4)) ? 42 + 5 * (k - 1) + j : `x${k}_${j}`;
const ACTION_COLS = new Set(Array.from({ length: 5 }, (_, i) => ACT.map((_, j) => ak(i + 1, j))).flat());
const acts = d => { let n = 0; for (let k = 1; k <= 99; k++) if (ACT.some((q, j) => d[ak(k, j)])) n = k; return n; };
const secsOf = d => S.secs.flatMap(s => s.loop ? Array.from({ length: acts(d) }, (_, i) => ({
  title: `SOLUCIÓN · ACCIÓN ${i + 1}`,
  qs: ACT.map((q, j) => ({ ...q, c: ak(i + 1, j), title: `${q.title} (acción ${i + 1})` }))
})) : [s]);

function fail(error) { if (error) throw new Error(error.message || String(error)); }
async function get(r) {
  const value = String(r || '').trim();
  let query = supabase.from('p').select('*');
  query = /^PQRS-\d+$/i.test(value) ? query.eq('codigo', value.toUpperCase()) : query.eq('id', value);
  const { data, error } = await query.maybeSingle();
  fail(error);
  if (!data) throw new Error('PQRS no encontrada');
  return data;
}
async function auth(r, c) {
  if (!E.CODIGO_EQUIPO || String(c || '').trim() !== String(E.CODIGO_EQUIPO).trim()) {
    throw new Error('Código incorrecto. Solo el equipo de Labtronics SAS puede continuar');
  }
  return get(r);
}
async function put(x) {
  const { error } = await supabase.from('p').update({ d: x.d, pos: x.pos, fin: x.fin }).eq('id', x.id);
  fail(error);
}
const link = x => `${BASE}/${x.codigo}`;

function cuerpo(x, nueva) {
  let t = nueva ? 'Buen dia,\n\nMe permito informar que se genero una nueva PQRS.\n\n'
    : 'Buen dia,\n\nMe permito informarle que la PQRS ha sido finalizada.\nA continuacion encontrara las respuestas registradas y, al final, el PDF de la PQRS.\n\n';
  t += `Con codigo: ${x.codigo}\n`;
  if (nueva) t += `\nContinua con la trazabilidad de la PQRS por medio de este enlace:\n${link(x)}\n\n`;
  secsOf(x.d).flatMap(s => s.qs).forEach(q => {
    const v = x.d[q.c]; if (q.t === 'k' || !v) return;
    t += `${q.title}:\n` + (q.t === 'f' ? 'Archivos adjuntos:\n' + v.split(', ').map(u => 'Adjunto: ' + u).join('\n') : v) + '\n\n';
  });
  return t;
}
function pdfHtml(x) {
  const body = secsOf(x.d).map(s => {
    const qs = s.qs.filter(q => q.t !== 'k' && !SKIP.test(q.title) && x.d[q.c]); if (!qs.length) return '';
    return `<div class=c><h2>${esc(s.title || 'Gastos')}</h2><div class=g>` + qs.map(q => `<div class=e><div class=k>${esc(q.title)}</div>` +
      (q.t === 'f' ? x.d[q.c].split(', ').map((u, i) => `<a href="${esc(u)}">Archivo ${i + 1}</a>`).join('<br>') : `<p>${esc(x.d[q.c]).replace(/\n/g, '<br>')}</p>`) + '</div>').join('') + '</div></div>';
  }).join('');
  return `<!doctype html><meta charset=utf-8><style>@page{size:A4;margin:6mm}body{font:9px Arial;color:#1b1f24}h1{font-size:17px;color:#1f3b5c;margin:0}
.m{width:100%;border-collapse:collapse;margin:8px 0}.m td{border:1px solid #d8e2ee;padding:3px 5px}.m .h{background:#f2f7ff;font-weight:bold;color:#1f3b5c;width:80px}
.c{border:1px solid #d8e2ee;border-radius:6px;margin:0 0 8px;overflow:hidden}h2{background:#d8e8ff;color:#0c3b76;font-size:10px;margin:0;padding:7px 10px;text-transform:uppercase}
.g{columns:2}.e{break-inside:avoid;padding:4px 6px}.k{background:#f8fbff;font-weight:bold;color:#244865;padding:3px}p{margin:2px 0}a{color:#1155cc}</style>
<h1>Cierre de PQRS</h1><div>Resumen completo de respuestas registradas · Código CO-FR-010 · Versión 4 · Fecha 2024-11-08</div>
<table class=m><tr><td class=h>Codigo</td><td>${esc(x.codigo)}</td><td class=h>Cliente</td><td>${esc(x.d[4] || 'Sin nombre')}</td></tr>
<tr><td class=h>Tipo de caso</td><td>${esc(x.d[10] || '')}</td><td class=h>ID respuesta</td><td>${esc(x.id)}</td></tr></table>${body}`;
}
async function storageUpload(key, buffer, contentType) {
  const { error } = await supabase.storage.from(BUCKET).upload(key, buffer, { contentType, upsert: true });
  fail(error);
}
async function storageDownload(key) {
  const { data, error } = await supabase.storage
    .from(BUCKET)
    .download(key);

  fail(error);

  if (!data) {
    throw new Error('Archivo no encontrado');
  }

  return {
    buffer: Buffer.from(await data.arrayBuffer()),
    contentType: data.type || 'application/octet-stream'
  };
}
function archivosDePQRS(x) {
  const archivos = new Set();

  // Identificar los archivos adjuntos guardados en las respuestas.
  const respuestas = x.d || {};

  for (const valor of Object.values(respuestas)) {
    if (typeof valor !== 'string') continue;

    const coincidencias = valor.matchAll(/\/f\/([^,\s]+)/g);

    for (const coincidencia of coincidencias) {
      try {
        const key = decodeURIComponent(coincidencia[1]);

        // Solo incluir archivos adjuntos generados por esta aplicación.
        if (key.startsWith('archivo-')) {
          archivos.add(key);
        }
      } catch (_) {
        // Ignorar enlaces que no se puedan interpretar.
      }
    }
  }

  // Incluir el PDF de cierre, si existe.
  archivos.add(`pdf-${x.id}.pdf`);

  return [...archivos];
}

async function eliminarArchivosPQRS(registros) {
  const archivos = [
    ...new Set(registros.flatMap(archivosDePQRS))
  ];

  // Supabase permite eliminar varios objetos por solicitud.
  for (let i = 0; i < archivos.length; i += 100) {
    const lote = archivos.slice(i, i + 100);

    const { error } = await supabase.storage
      .from(BUCKET)
      .remove(lote);

    fail(error);
  }
}
async function pdf(x) {
  const browser = await puppeteer.launch({ args: ['--no-sandbox'], executablePath: E.PUPPETEER_EXECUTABLE_PATH || undefined });
  try {
    const page = await browser.newPage();
    await page.setContent(pdfHtml(x));
    const buffer = Buffer.from(await page.pdf({ format: 'A4', printBackground: true }));
    await storageUpload(`pdf-${x.id}.pdf`, buffer, 'application/pdf');
    return buffer;
  } finally { await browser.close(); }
}

async function enviar(x, nueva, pdfBuffer) {
  const area = x.d[10] || '';
  const r = RUTAS[area] || [''];

  const destinatarios = [
    r[0],
    'asistentecomercial@labtronics.net'
  ]
    .filter(Boolean)
    .join(',')
    .split(',')
    .map(email => email.trim())
    .filter(Boolean)
    .map(email => ({ email }));

  const copias = (r[1] || '')
    .split(',')
    .map(email => email.trim())
    .filter(Boolean)
    .map(email => ({ email }));

  const asunto = nueva
    ? `Codigo PQRS ${x.codigo} de tipo: ${area || 'Sin tipo'}`
    : `Se finalizo la PQRS ${x.codigo}${x.d[4] ? ' - ' + x.d[4] : ''}`;

  let texto = cuerpo(x, nueva);

  const mensaje = {
    sender: {
      name: E.BREVO_SENDER_NAME || 'Labtronics SAS',
      email: E.BREVO_SENDER_EMAIL
    },
    to: destinatarios,
    subject: asunto,
    textContent: texto
  };

  if (copias.length) {
    mensaje.cc = copias;
  }

  if (!nueva && pdfBuffer) {
    mensaje.attachment = [{
      name: `Cierre_${x.codigo}.pdf`,
      content: pdfBuffer.toString('base64')
    }];

    mensaje.textContent +=
      `\nAl final encontrara el PDF de la PQRS en el siguiente enlace:\n${BASE}/pdf/${x.id}.pdf\n`;
  }

  if (!E.BREVO_API_KEY) {
    throw new Error('Falta configurar BREVO_API_KEY.');
  }

  if (!E.BREVO_SENDER_EMAIL) {
    throw new Error('Falta configurar BREVO_SENDER_EMAIL.');
  }

  const respuesta = await fetch(
    'https://api.brevo.com/v3/smtp/email',
    {
      method: 'POST',
      headers: {
        'accept': 'application/json',
        'api-key': E.BREVO_API_KEY,
        'content-type': 'application/json'
      },
      body: JSON.stringify(mensaje)
    }
  );

  if (!respuesta.ok) {
    const detalle = await respuesta.text();
    throw new Error(`Brevo respondió ${respuesta.status}: ${detalle}`);
  }

  const resultado = await respuesta.json();

  console.log('Correo enviado por Brevo:', resultado.messageId);

  return resultado;
}

async function filas() {
  const { data, error } = await supabase.from('p').select('*').order('num', { ascending: true });
  fail(error);
  return (data || []).map(x => {
    const d = x.d || {}, n = acts(d), g = (k, j) => d[ak(k, j)] || '';
    const solucion = Array.from({ length: n }, (_, i) => `${i + 1}) ${g(i + 1, 1)} · ${g(i + 1, 0)}: ${g(i + 1, 2)}${g(i + 1, 3) ? ' → ' + g(i + 1, 3) : ''}`).join('\n');
    return {
      id: x.id, no: x.codigo, fecha: x.ts, tipo: d[2] || d[10] || '',
      cliente: [4, 5, 6, 7, 8].map(c => d[c]).filter(Boolean).join('\n'),
      proceso: d[34] || '', descripcion: DESC.map(c => d[c]).find(Boolean) || '',
      contacto: [d[31], d[32]].filter(Boolean).join(' · '), investigacion: d[38] || d[40] || '',
      procedente: d[35] || '', solucion, fechaSol: d[75] || '', estado: x.fin ? 'Cerrado' : 'Abierto',
      cal: d.cal || '', acc: d.acc || '',
      dias: d[75] ? Math.max(0, Math.round((new Date(d[75]) - new Date(x.ts)) / 864e5)) : '',
      pdf: x.fin ? `${BASE}/pdf/${x.id}.pdf` : '', link: link(x)
    };
  });
}
async function excel() {
  const { data, error } = await supabase.from('p').select('*').order('num', { ascending: true });
  fail(error);
  const rows = (data || []).map(x => ({ x, d: x.d || {} }));
  const w = new ExcelJS.Workbook(), a = w.addWorksheet('Respuestas de formulario 1'), c = w.addWorksheet('CONSOLIDADO DE QUEJAS, RECLAMOS');
  const ex = [...new Set(rows.flatMap(r => Object.keys(r.d).filter(k => /^x\d+_\d$/.test(k))))].sort((p, q) => p.localeCompare(q, undefined, { numeric: true }));
  a.addRow([...H, ...ex.map(k => { const m = /x(\d+)_(\d)/.exec(k); return `Acción ${m[1]}: ${ACT[m[2]].title}`; })]);
  rows.forEach(({ x, d }) => {
    const r = new Array(H.length).fill(''); r[0] = new Date(x.ts);
    Object.keys(d).forEach(k => { if (/^\d+$/.test(k)) r[Number(k) - 1] = d[k]; });
    r[H.length - 1] = x.codigo;
    a.addRow([...r, ...ex.map(k => d[k] || '')]);
  });
  c.addRow(['CONSOLIDADO DE QUEJAS, RECLAMOS Y SUGERENCIAS']);
  c.addRow(['No', 'FECHA', 'TIPO', 'NOMBRE DEL CLIENTE', 'PROCESO RESPONSABLE', 'DESCRIPCIÓN DE LA PQRS', 'PRIMER CONTACTO CON EL CLIENTE', 'INVESTIGACIÓN / CAUSAS PROBABLES', 'PROCEDENTE', 'SOLUCIÓN DE LA PQRS (Acciones ejecutadas)', 'FECHA DE SOLUCIÓN', 'ESTADO', 'CALIFICACIÓN DE LA GESTIÓN', 'REQUIERE ACCIÓN', 'TIEMPO DE SOLUCIÓN (días)']);
  (await filas()).forEach(f => c.addRow([f.no, new Date(f.fecha), f.tipo, f.cliente, f.proceso, f.descripcion, f.contacto, f.investigacion, f.procedente, f.solucion, f.fechaSol, f.estado, f.cal, f.acc, f.dias]));
  c.columns.forEach(col => { col.width = 28; col.alignment = { wrapText: true, vertical: 'top' }; });
  return Buffer.from(await w.xlsx.writeBuffer());
}

const API = {
  api_crear: async d => {
    const { data: numero, error: rpcError } = await supabase.rpc('tomar_siguiente_numero');
    fail(rpcError);
    if (!Number.isInteger(numero) || numero < 1) throw new Error('No se pudo obtener el siguiente número de PQRS. Revise la configuración del contador.');
    const dd = {};
    Q.forEach(q => { if (q.p == 1 && d[q.c]) dd[q.c] = String(d[q.c]); });
    const x = { id: crypto.randomUUID(), num: numero, codigo: 'PQRS-' + String(numero).padStart(4, '0'), ts: new Date().toISOString(), pos: 0, fin: 0, d: dd };
    const { error } = await supabase.from('p').insert(x);
    fail(error);
    let errorCorreo = '';
    try { await enviar(x, true); } catch (e) { errorCorreo = String(e.message || e); console.error('Error enviando correo de nueva PQRS:', e); }
    return { r: x.id, codigo: x.codigo, url: link(x), errorCorreo };
  },
  api_entrar: async (r, c) => {
    const x = await auth(r, c);
    x.d = x.d || {};
    x.d[Q.find(q => q.t === 'k').c] = c;
    await put(x);
    return { v: x.d, url: link(x) };
  },
  api_guardar: async (r, c, col, v) => {
    const x = await auth(r, c);
    if (!(/^x\d{1,2}_[0-4]$/.test(col) || ACTION_COLS.has(col) || Q.some(q => q.c == col && q.p == 2 && q.t != 'k'))) throw new Error('Columna no permitida');
    x.d = x.d || {};
    if (v) x.d[col] = String(v); else delete x.d[col];
    await put(x);
    return 1;
  },
  api_subir: async (n, m, b) => {
    if (typeof b !== 'string' || b.length > 14e6) throw new Error('Archivo mayor a 10 MB');
    const id = `archivo-${crypto.randomUUID()}-${String(n).replace(/[^\w.-]/g, '_')}`;
    await storageUpload(id, Buffer.from(b, 'base64'), m || 'application/octet-stream');
    return `${BASE}/f/${encodeURIComponent(id)}`;
  },
  api_finalizar: async (r, c) => {
    const x = await auth(r, c);
    if (!secsOf(x.d || {}).flatMap(s => s.qs).some(q => /firma/i.test(q.title) && x.d[q.c])) return { ok: false, msg: 'Adjunte el archivo de su firma para finalizar la PQRS.' };
    const pdfBuffer = await pdf(x);
    x.fin = 1;
    await put(x);
    let errorCorreo = '';
    try { await enviar(x, false, pdfBuffer); } catch (e) { errorCorreo = String(e.message || e); console.error('Error enviando correo de cierre:', e); }
    return { ok: true, codigo: x.codigo, errorCorreo };
  }
};

const app = express();
app.use(express.json({ limit: '25mb' }));
const safeHeaders = res => {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('Content-Security-Policy', 'sandbox');
};
app.get('/f/:key', async (req, res) => {
  try {
    const key = decodeURIComponent(req.params.key);
    const archivo = await storageDownload(key);

    const tiposVisibles = [
      'application/pdf',
      'image/jpeg',
      'image/png',
      'image/gif',
      'image/webp',
      'text/plain'
    ];

    safeHeaders(res);

    res.setHeader('X-Content-Type-Options', 'nosniff');

    if (tiposVisibles.includes(archivo.contentType)) {
      res.setHeader('Content-Disposition', 'inline');
    } else {
      res.setHeader('Content-Disposition', 'attachment');
    }

    res.type(archivo.contentType).send(archivo.buffer);

  } catch (e) {
    console.error('Error al consultar archivo:', e);
    res.status(404).send('Archivo no encontrado');
  }
});
app.get('/pdf/:id.pdf', async (req, res) => {
  try {
    const x = await get(req.params.id);

    if (!x.fin) {
      return res.status(404).send('PDF no disponible');
    }

    const archivo = await storageDownload(`pdf-${x.id}.pdf`);

    safeHeaders(res);
    res.setHeader('Content-Disposition', 'inline');
    res.type('application/pdf').send(archivo.buffer);

  } catch (e) {
    console.error('Error al consultar PDF:', e);
    res.status(404).send('PDF no encontrado');
  }
});
app.get('/api/schema', (req, res) => res.json(S));
app.post('/api/:fn', async (req, res) => {
  try {
    const f = API[req.params.fn];
    if (!f) throw new Error('No existe');
    if (!Array.isArray(req.body)) throw new Error('Formato de solicitud no válido');
    res.json(await f(...req.body));
  } catch (e) {
    console.error(e);
    res.status(400).json({ error: e.message || 'Error interno' });
  }
});
const adm = (req, res, next) => {
  const header = req.headers.authorization || '';
  let credentials = '';
  try { credentials = header.startsWith('Basic ') ? Buffer.from(header.slice(6), 'base64').toString() : ''; } catch (_) {}
  const separator = credentials.indexOf(':');
  const user = separator >= 0 ? credentials.slice(0, separator) : '';
  const pass = separator >= 0 ? credentials.slice(separator + 1) : '';
  if (!E.ADMIN_PASS || user !== E.ADMIN_USER || pass !== E.ADMIN_PASS) return res.set('WWW-Authenticate', 'Basic realm="PQRS Admin"').status(401).end();
  next();
};
app.get('/admin', adm, (req, res) => res.sendFile(path.resolve('admin.html')));
app.get('/admin/data', adm, async (req, res) => { try { res.json(await filas()); } catch (e) { res.status(500).json({ error: e.message }); } });
// Eliminar una PQRS sin modificar el consecutivo.
app.post('/admin/delete', adm, async (req, res) => {
  try {
    const id = String(req.body.id || '').trim();

    if (!id) {
      return res.status(400).json({ error: 'Falta el identificador de la PQRS.' });
    }

    const registro = await get(id);

    // Primero eliminar los archivos relacionados.
    await eliminarArchivosPQRS([registro]);

    // Después eliminar el registro.
    const { error } = await supabase
      .from('p')
      .delete()
      .eq('id', registro.id);

    fail(error);

    res.json({ ok: true, mensaje: `Se eliminó ${registro.codigo}.` });

  } catch (e) {
    console.error('Error eliminando PQRS:', e);
    res.status(500).json({ error: e.message });
  }
});


// Eliminar todas las PQRS y reiniciar el consecutivo.
app.post('/admin/delete-all', adm, async (req, res) => {
  try {
    const confirmacion = String(req.body.confirmacion || '');

    if (confirmacion !== 'ELIMINAR') {
      return res.status(400).json({
        error: 'Confirmación no válida.'
      });
    }

    // Consultar todos los registros por páginas.
    const registros = [];
    let desde = 0;
    const tamano = 1000;

    while (true) {
      const { data, error } = await supabase
        .from('p')
        .select('id, d')
        .order('num', { ascending: true })
        .range(desde, desde + tamano - 1);

      fail(error);

      registros.push(...(data || []));

      if (!data || data.length < tamano) break;

      desde += tamano;
    }

    // Eliminar adjuntos y PDF relacionados.
    await eliminarArchivosPQRS(registros);

    // Eliminar todos los registros de PQRS.
    const { error: errorEliminar } = await supabase
      .from('p')
      .delete()
      .gte('num', 1);

    fail(errorEliminar);

    // Reiniciar el consecutivo para que la siguiente sea PQRS-0001.
    const { error: errorContador } = await supabase
      .from('configuracion_pqrs')
      .update({ siguiente_numero: 1 })
      .eq('id', true);

    fail(errorContador);

    res.json({
      ok: true,
      eliminadas: registros.length,
      siguiente_numero: 1,
      mensaje: 'Se eliminaron las PQRS y se reinició el consecutivo.'
    });

  } catch (e) {
    console.error('Error eliminando todas las PQRS:', e);
    res.status(500).json({ error: e.message });
  }
});
app.post('/admin/set', adm, async (req, res) => {
  try {
    const { id, k, v } = req.body;
    if (!['cal', 'acc'].includes(k)) return res.status(400).json({ error: 'Campo no permitido' });
    const x = await get(id); x.d = x.d || {};
    if (v) x.d[k] = String(v); else delete x.d[k];
    await put(x); res.json(1);
  } catch (e) { res.status(400).json({ error: e.message }); }
});
app.get('/admin/config', adm, async (req, res) => {
  try {
    const { data, error } = await supabase.from('configuracion_pqrs').select('siguiente_numero').eq('id', true).single();
    fail(error); res.json({ siguiente_numero: data.siguiente_numero });
  } catch (e) { res.status(500).json({ error: e.message }); }
});
app.post('/admin/config', adm, async (req, res) => {
  try {
    const numero = Number(req.body.siguiente_numero);
    if (!Number.isSafeInteger(numero) || numero < 1) return res.status(400).json({ error: 'Ingrese un número entero mayor que cero.' });
    const { data: maxRow, error: maxError } = await supabase.from('p').select('num').order('num', { ascending: false }).limit(1).maybeSingle();
    fail(maxError);
    if (maxRow && numero <= maxRow.num) return res.status(400).json({ error: `El próximo número debe ser mayor que ${maxRow.num}, que ya está registrado.` });
    const { error } = await supabase.from('configuracion_pqrs').update({ siguiente_numero: numero }).eq('id', true);
    fail(error);
    res.json({ ok: true, siguiente_numero: numero });
  } catch (e) { res.status(400).json({ error: e.message }); }
});
app.get('/admin/excel', adm, async (req, res) => {
  try { const buffer = await excel(); res.attachment('REPORTE DE PQRS (Respuestas).xlsx').type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet').send(buffer); }
  catch (e) { res.status(500).send(e.message); }
});
app.get(/^\/pqrs-\d+$/i, (req, res) => res.sendFile(path.resolve('public/index.html')));
app.use(express.static('public'));
app.listen(E.PORT || 3000, () => console.log('PQRS en puerto ' + (E.PORT || 3000)));
