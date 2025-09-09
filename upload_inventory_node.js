// upload_inventory_node.js
import 'dotenv/config';
import * as fs from 'fs';

// Import correcto de xlsx en ESM:
import * as XLSX from 'xlsx/xlsx.mjs';
XLSX.set_fs(fs);

import admin from 'firebase-admin';

// ===== Config desde .env =====
const SERVICE_ACCOUNT_PATH = process.env.SERVICE_ACCOUNT_PATH || './serviceAccount.json';
const FIREBASE_PROJECT_ID = process.env.FIREBASE_PROJECT_ID;
const EXCEL_PATH = process.env.EXCEL_PATH || './INVENTARIO CESREN.xlsx';
const SHEET_INDEX = Number(process.env.SHEET_INDEX ?? 0);
const COLLECTION_NAME = 'inventory';

// ===== Validaciones =====
if (!fs.existsSync(SERVICE_ACCOUNT_PATH)) {
  console.error(`❌ No se encontró SERVICE_ACCOUNT_PATH: ${SERVICE_ACCOUNT_PATH}`);
  process.exit(1);
}
if (!FIREBASE_PROJECT_ID) {
  console.error('❌ Falta FIREBASE_PROJECT_ID en .env');
  process.exit(1);
}
if (!fs.existsSync(EXCEL_PATH)) {
  console.error(`❌ No se encontró el Excel en: ${EXCEL_PATH}`);
  process.exit(1);
}

// ===== Inicializa Admin SDK =====
const serviceAccount = JSON.parse(fs.readFileSync(SERVICE_ACCOUNT_PATH, 'utf-8'));
admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
  projectId: FIREBASE_PROJECT_ID,
});
const db = admin.firestore();

// ===== Normalización/mapeo de columnas =====
// Ajusta las llaves de la IZQUIERDA a como vienen en TU Excel y las de la DERECHA
// a como quieres verlas en Firestore. (Sensible a mayúsculas/minúsculas.)
const MAPPING = {
  // Ejemplos; ajusta según tus encabezados reales:
  // "Descripción": "descripcion",
  // "Referencia": "referencia",
  // "Marca": "marca",
  // "Cantidad": "cantidad",
  // "Estado": "estado",
  // "Fecha de registro": "fecharegistro",
  // "Fecha Registro": "fecharegistro",
  // "fecharegistro": "fecharegistro",
  // "Caducidad": "caducidad",
  // "Fecha de caducidad": "caducidad",
  // "caducidad": "caducidad",
};

// Si quieres usar un ID personalizado por fila (p. ej. "referencia"), pon el campo destino:
const CUSTOM_ID_FIELD = ''; // por ejemplo: 'referencia'

// ===== Utilidades de fecha =====
/**
 * Convierte cualquier valor (serial Excel, Date, string) a "YYYY-MM-DD".
 * Si no puede interpretarse, devuelve la cadena original "tal cual".
 */
function toYYYYMMDD(value) {
  if (value == null || value === '') return ''; // vacío

  // 1) Si ya es Date
  if (value instanceof Date && !isNaN(value)) {
    return value.toISOString().slice(0, 10);
  }

  // 2) Si es número: asumir serial de Excel (sistema 1900)
  if (typeof value === 'number' && !isNaN(value)) {
    // Excel serial -> JS Date
    // 25569 = días entre 1899-12-30 y 1970-01-01
    const ms = Math.round((value - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if (!isNaN(d)) return d.toISOString().slice(0, 10);
  }

  // 3) Si es string: intentar parsear
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return '';

    // Intentos comunes:
    // a) ISO (YYYY-MM-DD o YYYY-MM-DDTHH:mm:ss)
    const isoCandidate = new Date(trimmed);
    if (!isNaN(isoCandidate)) return isoCandidate.toISOString().slice(0, 10);

    // b) Formatos dd/mm/yyyy o dd-mm-yyyy
    const m = trimmed.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (m) {
      let [_, dd, mm, yyyy] = m;
      if (yyyy.length === 2) yyyy = `20${yyyy}`; // naive: 24 -> 2024
      const d = new Date(Number(yyyy), Number(mm) - 1, Number(dd));
      if (!isNaN(d)) return d.toISOString().slice(0, 10);
    }

    // No se pudo parsear: devolver tal cual (pero sin hora)
    // Si viene con hora "YYYY-MM-DD HH:MM", intenta separar:
    const cut = trimmed.split(' ')[0];
    if (/^\d{4}\-\d{2}\-\d{2}$/.test(cut)) return cut;

    return trimmed; // fallback
  }

  // Cualquier otro tipo
  return String(value);
}

const cleanValue = (v) =>
  v === null || v === undefined || (typeof v === 'number' && Number.isNaN(v)) ? '' : v;

/**
 * Construye el objeto destino aplicando MAPPING y formato de fechas.
 * Solo formatea 'fecharegistro' y 'caducidad'.
 */
function buildDocData(rawRow) {
  const docData = {};
  for (const key of Object.keys(rawRow)) {
    const originalKey = String(key).trim();
    const targetKey = MAPPING[originalKey] || originalKey; // si no hay mapping, deja como está
    let val = rawRow[originalKey];

    if (targetKey.toLowerCase() === 'fecharegistro') {
      val = toYYYYMMDD(val);
    } else if (targetKey.toLowerCase() === 'caducidad') {
      val = toYYYYMMDD(val);
    } else {
      val = cleanValue(val);
    }

    docData[targetKey] = val;
  }
  return docData;
}

// ===== Lectura del Excel =====
function readExcelRows(path, sheetIndex = 0) {
  const workbook = XLSX.readFile(path);
  const sheetName = workbook.SheetNames[sheetIndex];
  if (!sheetName) throw new Error(`No existe la hoja con índice ${sheetIndex}.`);
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

// ===== Main =====
async function main() {
  console.log('📄 Leyendo Excel…');
  const rows = readExcelRows(EXCEL_PATH, SHEET_INDEX);
  console.log(`✅ Filas encontradas: ${rows.length}`);

  const batchSize = 400;
  let ok = 0, fail = 0;

  for (let start = 0; start < rows.length; start += batchSize) {
    const slice = rows.slice(start, start + batchSize);
    const batch = db.batch();

    for (const rawRow of slice) {
      try {
        const docData = buildDocData(rawRow);

        if (CUSTOM_ID_FIELD && docData[CUSTOM_ID_FIELD]) {
          const id = String(docData[CUSTOM_ID_FIELD]).trim();
          if (!id) throw new Error(`ID vacío en campo ${CUSTOM_ID_FIELD}`);
          const ref = db.collection(COLLECTION_NAME).doc(id);
          batch.set(ref, docData, { merge: true });
        } else {
          const ref = db.collection(COLLECTION_NAME).doc();
          batch.set(ref, docData);
        }
        ok++;
      } catch (e) {
        fail++;
        console.error('Error preparando doc:', e.message, 'Fila:', rawRow);
      }
    }

    await batch.commit();
    console.log(`🔹 Subidas: ${Math.min(start + batchSize, rows.length)}/${rows.length}`);
  }

  console.log(`\n🎉 Listo. Documentos OK: ${ok}, Errores: ${fail}`);
  console.log(`Colección: ${COLLECTION_NAME}`);
}

main().catch((e) => {
  console.error('❌ Error general:', e);
  process.exit(1);
});
