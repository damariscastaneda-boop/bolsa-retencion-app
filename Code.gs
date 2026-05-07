// ─────────────────────────────────────────────────────────────────────────────
// BIA Energy · Gestión de Bolsa de Retención
// ─────────────────────────────────────────────────────────────────────────────

const SPREADSHEET_ID  = '1EKVPE7q7sxoP7yKJVW3IRftHULQ9QXGa-3h6iX2L0bg';
const DATA_SHEET_NAME = 'BD_Clientes';
const CASOS_SHEET     = 'Bolsa_Casos';
const BOLSA_MENSUAL   = 15000000;     // $15.000.000 desde mayo 2026
const EXECUTIVE_PIN   = '2026';       // cámbialo por el PIN que quieras

const CONCIERGES = [
  'Andrés', 'Cristian', 'Cristina', 'Estefany',
  'Jerenny', 'Lina', 'Susan', 'Vanessa'
];

const MOTIVOS = [
  'Tarifas / Competitividad',
  'Emergencias',
  'Calidad de la energía / Reactivas',
  'Pagos / Plazos / AGPE',
  'Facturación / SAP',
  'Impuestos / Alumbrado',
  'Upselling no atendido',
  'Ingreso EMS',
  'Financiación',
  'Otro'
];

// ─── WEB APP ─────────────────────────────────────────────────────────────────

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('BIA · Gestión de Retención')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── BÚSQUEDA DE CLIENTES ────────────────────────────────────────────────────

function searchClients(query) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DATA_SHEET_NAME) || ss.getSheets()[0];
    const data  = sheet.getDataRange().getValues();
    const hdrs  = data[0].map(h => String(h).trim());

    const C = {
      compania:         hdrs.indexOf('RAZÓN SOCIAL'),
      nivel:            hdrs.indexOf('Nivel Servicio'),
      compania_id:      hdrs.indexOf('Company ID'),
      nit:              hdrs.indexOf('NIT'),
      active_energy:    hdrs.indexOf('Consumo último Mes'),
      total_cb:         hdrs.indexOf('C+B ÚLTIMO MES'),
      market_type:      hdrs.indexOf('Mercado'),
      gmv:              hdrs.indexOf('GMV'),
      city:             hdrs.indexOf('Ciudad'),
      fecha_activacion: hdrs.indexOf('Fecha activación'),
    };

    const q = query.toLowerCase().trim();
    const map = {};

    for (let i = 1; i < data.length; i++) {
      const row  = data[i];
      const name = String(row[C.compania] || '').trim();
      const cid  = String(row[C.compania_id] || '').trim();
      const nit  = String(row[C.nit] || '').trim();
      if (!name) continue;
      if (!name.toLowerCase().includes(q) && !cid.includes(q) && !nit.includes(q)) continue;

      if (!map[cid]) {
        map[cid] = {
          compania:         name,
          compania_id:      cid,
          nit:              nit,
          nivel:            String(row[C.nivel] || '').trim(),
          market_type:      String(row[C.market_type] || '').trim(),
          city:             String(row[C.city] || '').trim(),
          fecha_activacion: String(row[C.fecha_activacion] || '').trim(),
          total_energy:     0,
          total_cb:         0,
          total_gmv:        0,
          contracts:        0,
        };
      }
      map[cid].total_energy += parseNum(row[C.active_energy]);
      map[cid].total_cb     += parseNum(row[C.total_cb]);
      map[cid].total_gmv    += parseNum(row[C.gmv]);
      map[cid].contracts++;
    }

    const results = Object.values(map).map(c => {
      const take   = c.total_gmv > 0 ? (c.total_cb / c.total_gmv) * 100 : 0;
      const months = monthsWithBia(c.fecha_activacion);
      return {
        compania:        c.compania,
        compania_id:     c.compania_id,
        nit:             c.nit,
        nivel:           c.nivel,
        kwh:             Math.round(c.total_energy),
        cb:              Math.round(c.total_cb),
        gmv:             Math.round(c.total_gmv),
        take_rate:       take.toFixed(1),
        mercado:         c.market_type === 'REGULATED' ? 'Regulado' : 'No Regulado',
        ciudad:          c.city,
        meses_con_bia:   months,
        momento_default: months >= 12
          ? 'Fricción relación (Lleva ≥ 1 año)'
          : 'Fricción relación (Lleva < 1 año)',
        contracts:       c.contracts,
      };
    });

    return { ok: true, data: results.slice(0, 12) };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function parseNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const s = String(val).replace(/\./g, '').replace(',', '.').replace(/[^0-9.-]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function monthsWithBia(fechaStr) {
  if (!fechaStr) return 0;
  try {
    const d = new Date(fechaStr);
    if (isNaN(d.getTime())) return 0;
    const now = new Date();
    return Math.max(0, (now.getFullYear() - d.getFullYear()) * 12 + (now.getMonth() - d.getMonth()));
  } catch (_) { return 0; }
}

// ─── GESTIÓN DE CASOS ────────────────────────────────────────────────────────

function getCasosSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh   = ss.getSheetByName(CASOS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CASOS_SHEET);
    const headers = [
      'id','fecha_solicitud','compania_id','compania','nivel',
      'kwh','gmv','cb','take_rate','mercado','ciudad','meses_con_bia',
      'concierge','motivo','momento','descripcion',
      'estado','fecha_decision','que_hicimos','valor_bolsa','mes_entrega'
    ];
    sh.appendRow(headers);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#472bef').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function submitCase(d) {
  try {
    const sh  = getCasosSheet();
    const id  = 'RET-' + Date.now();
    const now = new Date().toISOString();
    sh.appendRow([
      id, now,
      d.compania_id, d.compania, d.nivel,
      d.kwh, d.gmv, d.cb, d.take_rate,
      d.mercado, d.ciudad, d.meses_con_bia,
      d.concierge, d.motivo, d.momento, d.descripcion,
      'PENDIENTE', '', '', '', ''
    ]);
    return { ok: true, id };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function getPendingCases() {
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh   = ss.getSheetByName(CASOS_SHEET);
    if (!sh) return { ok: true, data: [] };
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return { ok: true, data: [] };
    const hdrs = data[0];
    return {
      ok: true,
      data: data.slice(1)
        .filter(r => String(r[16]).trim().toUpperCase() === 'PENDIENTE')
        .map(r => toObj(hdrs, r))
        .reverse()
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function getAllExecData() {
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh   = ss.getSheetByName(CASOS_SHEET);
    const now  = new Date();
    const mes  = now.getMonth();
    const anio = now.getFullYear();

    if (!sh) return {
      ok: true,
      stats: { bolsa_total: BOLSA_MENSUAL, usado: 0, disponible: BOLSA_MENSUAL, pct_usado: 0, casos_mes: 0, casos_sin_costo: 0 },
      pending: [],
      historial: []
    };

    const data = sh.getDataRange().getValues();
    const hdrs = data[0];
    const rows = data.slice(1);

    let usado = 0, casosAprobados = 0, casosSinCosto = 0;
    rows.forEach(r => {
      if (String(r[16]).trim() !== 'APROBADO') return;
      const fd = new Date(r[17]);
      if (fd.getMonth() !== mes || fd.getFullYear() !== anio) return;
      const val = parseFloat(r[19]) || 0;
      usado += val;
      casosAprobados++;
      if (val === 0) casosSinCosto++;
    });

    const pending  = rows.filter(r => String(r[16]).trim().toUpperCase() === 'PENDIENTE').map(r => toObj(hdrs, r)).reverse();
    const historial = rows.filter(r => { const s = String(r[16]).trim(); return s === 'APROBADO' || s === 'RECHAZADO'; }).map(r => toObj(hdrs, r)).reverse();

    return {
      ok: true,
      stats: {
        bolsa_total: BOLSA_MENSUAL,
        usado, disponible: BOLSA_MENSUAL - usado,
        pct_usado: +((usado / BOLSA_MENSUAL) * 100).toFixed(1),
        casos_mes: casosAprobados,
        casos_sin_costo: casosSinCosto
      },
      pending,
      historial
    };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function approveCase(id, det) {
  try {
    const sh   = getCasosSheet();
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sh.getRange(i+1,17).setValue('APROBADO');
        sh.getRange(i+1,18).setValue(new Date().toISOString());
        sh.getRange(i+1,19).setValue(det.que_hicimos || '');
        sh.getRange(i+1,20).setValue(parseFloat(det.valor_bolsa) || 0);
        sh.getRange(i+1,21).setValue(det.mes_entrega || '');
        return { ok: true };
      }
    }
    return { ok: false, error: 'Caso no encontrado' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function rejectCase(id) {
  try {
    const sh   = getCasosSheet();
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sh.getRange(i+1,17).setValue('RECHAZADO');
        sh.getRange(i+1,18).setValue(new Date().toISOString());
        return { ok: true };
      }
    }
    return { ok: false, error: 'Caso no encontrado' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function getBolsaStats() {
  try {
    const sh   = getCasosSheet();
    const data = sh.getDataRange().getValues();
    const now  = new Date();
    const mes  = now.getMonth();
    const anio = now.getFullYear();

    let usado = 0, casosAprobados = 0, casosSinCosto = 0;

    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (r[16] !== 'APROBADO') continue;
      const fd = new Date(r[17]);
      if (fd.getMonth() !== mes || fd.getFullYear() !== anio) continue;
      const val = parseFloat(r[19]) || 0;
      usado += val;
      casosAprobados++;
      if (val === 0) casosSinCosto++;
    }

    return {
      ok: true,
      data: {
        bolsa_total:     BOLSA_MENSUAL,
        usado:           usado,
        disponible:      BOLSA_MENSUAL - usado,
        pct_usado:       +((usado / BOLSA_MENSUAL) * 100).toFixed(1),
        casos_mes:       casosAprobados,
        casos_sin_costo: casosSinCosto,
      }
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function getHistorial() {
  try {
    const sh   = getCasosSheet();
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return { ok: true, data: [] };
    const hdrs = data[0];
    return {
      ok: true,
      data: data.slice(1)
        .filter(r => r[16] === 'APROBADO' || r[16] === 'RECHAZADO')
        .map(r => toObj(hdrs, r))
        .reverse()
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function getConfig() {
  return {
    ok: true,
    concierges: CONCIERGES,
    motivos: MOTIVOS,
  };
}

function verifyPin(pin) {
  return { ok: String(pin) === String(EXECUTIVE_PIN) };
}

// ─── UTILS ───────────────────────────────────────────────────────────────────

function toObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => { obj[h] = row[i]; });
  return obj;
}

function debugSheet() {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets().map(s => s.getName());
  const sheet  = ss.getSheets()[0];
  const hdrs   = sheet.getRange(1, 1, 1, 25).getValues()[0];
  Logger.log('Pestañas: ' + JSON.stringify(sheets));
  Logger.log('Columnas fila 1: ' + JSON.stringify(hdrs));
  Logger.log('Fila 2 (primer dato): ' + JSON.stringify(sheet.getRange(2, 1, 1, 25).getValues()[0]));
}
