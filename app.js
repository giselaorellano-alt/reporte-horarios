/**
 * Lógica principal.
 * No hace falta tocar este archivo para cambiar el orden de columnas:
 * eso se edita en config.js.
 *
 * Usa ExcelJS (no SheetJS) porque necesitamos que el Excel de salida
 * mantenga fuente, colores de fondo, bordes e hipervínculos del original,
 * y la versión gratuita de SheetJS no escribe estilos.
 */

const dropzone = document.getElementById("dropzone");
const dropzoneText = document.getElementById("dropzoneText");
const fileInput = document.getElementById("fileInput");
const statusEl = document.getElementById("status");
const downloadBtn = document.getElementById("downloadBtn");

let generatedBlob = null;
let generatedFileName = "";

// --- Eventos de drag & drop / selección de archivo -------------------

["dragover", "dragenter"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  })
);

["dragleave", "drop"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
  })
);

dropzone.addEventListener("drop", (e) => {
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});

fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (file) handleFile(file);
});

downloadBtn.addEventListener("click", () => {
  if (!generatedBlob) return;
  const url = URL.createObjectURL(generatedBlob);
  const a = document.createElement("a");
  a.href = url;
  a.download = generatedFileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
});

// --- Procesamiento del archivo -----------------------------------------

function setStatus(message, type) {
  statusEl.textContent = message;
  statusEl.className = "status" + (type ? " " + type : "");
}

async function handleFile(file) {
  downloadBtn.disabled = true;
  generatedBlob = null;
  dropzoneText.textContent = file.name;
  setStatus("Procesando...", "");

  try {
    const arrayBuffer = await file.arrayBuffer();
    const sourceWb = new ExcelJS.Workbook();
    await sourceWb.xlsx.load(arrayBuffer);
    const sourceWs = sourceWb.worksheets[0];

    const outWb = new ExcelJS.Workbook();
    const outWs = outWb.addWorksheet(sourceWs.name || "Reporte");

    const result = reorderSheet(sourceWs, outWs);

    const buffer = await outWb.xlsx.writeBuffer();
    generatedBlob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const baseName = file.name.replace(/\.(xlsx|xls)$/i, "");
    const stamp = new Date().toISOString().slice(0, 10);
    generatedFileName = `${OUTPUT_FILE_PREFIX}_${baseName}_${stamp}.xlsx`;

    downloadBtn.disabled = false;
    setStatus(
      `Listo. ${result.rowCount} filas de datos, ${
        result.missingColumns.length
          ? "columnas sin dato en el original: " + result.missingColumns.join(", ")
          : "todas las columnas encontradas"
      }.`,
      "ok"
    );
  } catch (err) {
    console.error(err);
    setStatus("Error al procesar el archivo: " + err.message, "error");
  }
}

/**
 * Busca en las primeras filas cuál es la fila de encabezados
 * (la que contiene la celda "Usuario"), sin asumir que siempre
 * es la fila 2 como en el ejemplo.
 */
function findHeaderRow(ws) {
  const maxScan = Math.min(10, ws.rowCount);
  for (let r = 1; r <= maxScan; r++) {
    const row = ws.getRow(r);
    let found = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (typeof cell.value === "string" && cell.value.trim().toLowerCase() === "usuario") {
        found = true;
      }
    });
    if (found) return r;
  }
  throw new Error('No encontré la fila de encabezados (celda "Usuario").');
}

/** Copia value + estilo completo (fuente, fondo, borde, alineación, formato numérico). */
function cloneCellInto(sourceCell, destCell) {
  if (sourceCell.value !== null && sourceCell.value !== undefined) {
    destCell.value = sourceCell.value;
  }
  destCell.style = JSON.parse(JSON.stringify(sourceCell.style || {}));
  if (sourceCell.numFmt) destCell.numFmt = sourceCell.numFmt;
}

/**
 * Reordena las columnas de la hoja de entrada según TARGET_COLUMNS,
 * escribiendo directamente sobre outWs. Devuelve info de columnas
 * que no existían en el original.
 */
function reorderSheet(ws, outWs) {
  const headerRow = findHeaderRow(ws);

  // Mapa: nombre de columna (normalizado) -> número de columna original (1-based)
  const sourceColByName = {};
  ws.getRow(headerRow).eachCell({ includeEmpty: false }, (cell, colNumber) => {
    if (typeof cell.value === "string" && cell.value.trim() !== "") {
      sourceColByName[cell.value.trim().toLowerCase()] = colNumber;
    }
  });

  const missingColumns = TARGET_COLUMNS.filter(
    (name) => !(name.trim().toLowerCase() in sourceColByName)
  );

  // --- Fila(s) de título (todo lo que está antes del header) ---
  const headerStyleSample = ws.getRow(headerRow).getCell(sourceColByName["usuario"]);

  for (let r = 1; r < headerRow; r++) {
    const sourceRow = ws.getRow(r);
    const titleCell = sourceRow.getCell(1);
    if (titleCell.value) {
      cloneCellInto(titleCell, outWs.getCell(r, 1));
    }
    // Reubicar "Período" y su valor en la columna de "Fecha" del nuevo orden
    sourceRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (typeof cell.value === "string" && cell.value.trim().toLowerCase() === "período") {
        const valueCell = sourceRow.getCell(colNumber + 1);
        const fechaIdx = TARGET_COLUMNS.findIndex((n) => n === "Fecha");
        const targetCol = (fechaIdx >= 0 ? fechaIdx : 5) + 1;
        cloneCellInto(cell, outWs.getCell(r, targetCol));
        if (valueCell) cloneCellInto(valueCell, outWs.getCell(r, targetCol + 1));
      }
    });
  }
  // Merge visual del título, igual que en el original (si existía)
  const originalMerge = (ws.model.merges || [])[0];
  if (originalMerge) {
    try {
      outWs.mergeCells(originalMerge);
    } catch (e) {
      /* si no se puede mergear, no rompe el resto del proceso */
    }
  }

  // --- Fila de encabezados nueva, con el orden de TARGET_COLUMNS ---
  TARGET_COLUMNS.forEach((name, idx) => {
    const destCell = outWs.getCell(headerRow, idx + 1);
    destCell.value = name;
    if (headerStyleSample) {
      destCell.style = JSON.parse(JSON.stringify(headerStyleSample.style || {}));
    }
  });

  // --- Filas de datos ---
  let rowCount = 0;
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const sourceRow = ws.getRow(r);
    let rowHasData = false;
    sourceRow.eachCell({ includeEmpty: false }, (cell) => {
      if (cell.value !== null && cell.value !== undefined && cell.value !== "") rowHasData = true;
    });
    if (!rowHasData) continue;

    const outR = headerRow + 1 + rowCount;
    TARGET_COLUMNS.forEach((name, idx) => {
      const sourceCol = sourceColByName[name.trim().toLowerCase()];
      if (sourceCol === undefined) return; // columna no existe en el original -> queda en blanco
      const sourceCell = sourceRow.getCell(sourceCol);
      if (sourceCell.value === null || sourceCell.value === undefined) return;
      cloneCellInto(sourceCell, outWs.getCell(outR, idx + 1));
    });
    rowCount++;
  }

  // --- Ancho de columnas (copiado del original cuando existe) ---
  outWs.columns = TARGET_COLUMNS.map((name) => {
    const sourceCol = sourceColByName[name.trim().toLowerCase()];
    const sourceWidth = sourceCol ? ws.getColumn(sourceCol).width : null;
    return { width: sourceWidth || Math.max(12, name.length + 2) };
  });

  return { rowCount, missingColumns };
}

// Export para poder testear las funciones puras desde Node (no afecta el navegador)
if (typeof module !== "undefined") {
  module.exports = { findHeaderRow, reorderSheet, cloneCellInto };
}
