// AQUAELITE | Historial editable de cotizaciones
// Integrar estas funciones al Apps Script del Portal Unificado.

const QUOTE_VERSIONS_SHEET_NAME = "COTIZACION_VERSIONES";

function ensureQuoteVersionsSheet_() {
  const ss = SpreadsheetApp.openById(CRM_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(QUOTE_VERSIONS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(QUOTE_VERSIONS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 10).setValues([[
      "N° COTIZACIÓN", "REVISIÓN", "FECHA GUARDADO", "TIPO", "CLIENTE",
      "ESTADO", "PDF / ARCHIVO", "JSON COTIZACIÓN", "EDITADO POR", "OBSERVACIONES"
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveQuoteState(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const numero = clean(data.numero || "");
    const payload = String(data.payload || "").trim();
    if (!numero) return { ok: false, message: "Falta el número de cotización." };
    if (!payload) return { ok: false, message: "Falta el contenido editable de la cotización." };
    try { JSON.parse(payload); } catch (e) {
      return { ok: false, message: "El contenido editable no es JSON válido." };
    }

    const sheet = ensureQuoteVersionsSheet_();
    const lastRow = sheet.getLastRow();
    let revision = 0;
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      values.forEach(function(row) {
        if (String(row[0]).trim() === numero) {
          const r = Number(row[1]) || 0;
          if (r > revision) revision = r;
        }
      });
    }
    revision += 1;

    sheet.appendRow([
      numero,
      revision,
      new Date(),
      clean(data.tipo || ""),
      clean(data.cliente || ""),
      clean(data.estado || "BORRADOR"),
      clean(data.pdf_nombre || ""),
      payload,
      clean(data.editor || "Cotizador Aquaelite"),
      clean(data.observaciones || "")
    ]);

    return {
      ok: true,
      numero: numero,
      revision: revision,
      row: sheet.getLastRow(),
      message: "Versión editable guardada correctamente."
    };
  } catch (err) {
    return { ok: false, message: "No se pudo guardar la versión editable.", error: String(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function loadQuoteState(data) {
  try {
    const numero = clean(data.numero || "");
    if (!numero) return { ok: false, message: "Falta el número de cotización." };
    const requestedRevision = Number(data.revision || 0);
    const sheet = ensureQuoteVersionsSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: false, message: "No existen versiones guardadas." };

    const rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const matches = rows.map(function(row, index) {
      return { row: index + 2, values: row, revision: Number(row[1]) || 0 };
    }).filter(function(item) {
      return String(item.values[0]).trim() === numero && (!requestedRevision || item.revision === requestedRevision);
    });

    if (!matches.length) return { ok: false, message: "No se encontró la cotización o revisión solicitada." };
    matches.sort(function(a, b) { return b.revision - a.revision; });
    const found = matches[0].values;

    return {
      ok: true,
      numero: String(found[0]),
      revision: Number(found[1]) || 0,
      fecha_guardado: found[2] instanceof Date ? found[2].toISOString() : String(found[2] || ""),
      tipo: String(found[3] || ""),
      cliente: String(found[4] || ""),
      estado: String(found[5] || ""),
      pdf_nombre: String(found[6] || ""),
      payload: String(found[7] || ""),
      editor: String(found[8] || ""),
      observaciones: String(found[9] || "")
    };
  } catch (err) {
    return { ok: false, message: "No se pudo abrir la cotización guardada.", error: String(err) };
  }
}

function listQuoteVersions(data) {
  try {
    const numero = clean(data.numero || "");
    if (!numero) return { ok: false, message: "Falta el número de cotización." };
    const sheet = ensureQuoteVersionsSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, numero: numero, versions: [] };

    const rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const versions = rows.filter(function(row) {
      return String(row[0]).trim() === numero;
    }).map(function(row) {
      return {
        revision: Number(row[1]) || 0,
        fecha_guardado: row[2] instanceof Date ? row[2].toISOString() : String(row[2] || ""),
        estado: String(row[5] || ""),
        pdf_nombre: String(row[6] || ""),
        editor: String(row[8] || "")
      };
    }).sort(function(a, b) { return b.revision - a.revision; });

    return { ok: true, numero: numero, versions: versions };
  } catch (err) {
    return { ok: false, message: "No se pudo listar el historial de la cotización.", error: String(err) };
  }
}
