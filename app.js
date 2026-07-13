/**
 * Lógica principal.
 * No hace falta tocar este archivo para cambiar el orden de columnas:
 * eso se edita en config.js.
 */

const dropzone = document.getElementById("dropzone");
const dropzoneText = document.getElementById("dropzoneText");
const fileInput = document.getElementById("fileInput");
const statusEl = document.getElementById("status");
const downloadBtn = document.getElementById("downloadBtn");

let generatedWorkbook = null;
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
  if (!generatedWorkbook) return;
  XLSX.writeFile(generatedWorkbook, generatedFileName);
});

// --- Procesamiento del archivo -----------------------------------------

function setStatus(message, type) {
  statusEl.textContent = message;
  statusEl.className = "status" + (type ? " " + type : "");
}

function handleFile(file) {
  downloadBtn.disabled = true;
  generatedWorkbook = null;
  dropzoneText.textContent = file.name;
  setStatus("Procesando...", "");

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const sourceWb = XLSX.read(data, { type: "array", cellDates: false, cellNF: true });
      const sourceSheetName = sourceWb.SheetNames[0];
      const sourceWs = sourceWb.Sheets[sourceSheetName];

      const result = reorderSheet(sourceWs);

      const outWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(outWb, result.sheet, sourceSheetName);

      generatedWorkbook = outWb;
      const baseName = file.name.replace(/\.(xlsx|xls)$/i, "");
      const stamp = new Date().toISOString().slice(0, 10);
      generatedFileName = `${OUTPUT_FILE_PREFIX}_${baseName}_${stamp}.xlsx`;

      downloadBtn.disabled = false;
      setStatus(
        `Listo. ${result.rowCount} filas de datos, ${result.missingColumns.length ? "columnas sin dato en el original: " + result.missingColumns.join(", ") : "todas las columnas encontradas"}.`,
        "ok"
      );
    } catch (err) {
      console.error(err);
      setStatus("Error al procesar el archivo: " + err.message, "error");
    }
  };
  reader.onerror = () => setStatus("No se pudo leer el archivo.", "error");
  reader.readAsArrayBuffer(file);
}

/**
 * Busca en las primeras filas cuál es la fila de encabezados
 * (la que contiene la celda "Usuario"), sin asumir que siempre
 * es la fila 2 como en el ejemplo.
 */
function findHeaderRow(ws, range) {
  for (let r = range.s.r; r <= Math.min(range.s.r + 10, range.e.r); r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === "string" && cell.v.trim().toLowerCase() === "usuario") {
        return r;
      }
    }
  }
  throw new Error('No encontré la fila de encabezados (celda "Usuario").');
}

/**
 * Reordena las columnas de la hoja de entrada según TARGET_COLUMNS.
 * Devuelve la nueva hoja + info de columnas que no existían en el original.
 */
function reorderSheet(ws) {
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const headerRow = findHeaderRow(ws, range);

  // Mapa: nombre de columna (normalizado) -> índice de columna original
  const sourceColByName = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: headerRow, c })];
    if (cell && typeof cell.v === "string" && cell.v.trim() !== "") {
      sourceColByName[cell.v.trim().toLowerCase()] = c;
    }
  }

  const missingColumns = TARGET_COLUMNS.filter(
    (name) => !(name.trim().toLowerCase() in sourceColByName)
  );

  const newWs = {};
  let maxRow = 0;

  // Fila de título (todo lo que estaba antes del header, ej: "Reporte..." / "Período")
  for (let r = range.s.r; r < headerRow; r++) {
    const titleCell = ws[XLSX.utils.encode_cell({ r, c: range.s.c })];
    if (titleCell) {
      newWs[XLSX.utils.encode_cell({ r: r, c: 0 })] = { t: "s", v: titleCell.v || REPORT_TITLE };
    }
    // Buscar "Período" y su valor en la fila original para reubicarlo
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === "string" && cell.v.trim().toLowerCase() === "período") {
        const valueCell = ws[XLSX.utils.encode_cell({ r, c: c + 1 })];
        const periodoCol = TARGET_COLUMNS.findIndex((n) => n === "Fecha");
        const col = periodoCol >= 0 ? periodoCol : 5;
        newWs[XLSX.utils.encode_cell({ r, c: col })] = { t: "s", v: "Período" };
        if (valueCell) {
          newWs[XLSX.utils.encode_cell({ r, c: col + 1 })] = { t: "s", v: String(valueCell.v) };
        }
      }
    }
    maxRow = Math.max(maxRow, r);
  }

  // Fila de encabezados nueva, con el orden de TARGET_COLUMNS
  const newHeaderRow = headerRow;
  TARGET_COLUMNS.forEach((name, idx) => {
    newWs[XLSX.utils.encode_cell({ r: newHeaderRow, c: idx })] = { t: "s", v: name };
  });
  maxRow = Math.max(maxRow, newHeaderRow);

  // Filas de datos
  let rowCount = 0;
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    // saltear filas completamente vacías
    let rowHasData = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      if (ws[XLSX.utils.encode_cell({ r, c })]) {
        rowHasData = true;
        break;
      }
    }
    if (!rowHasData) continue;

    const outR = headerRow + 1 + rowCount;
    TARGET_COLUMNS.forEach((name, idx) => {
      const sourceCol = sourceColByName[name.trim().toLowerCase()];
      if (sourceCol === undefined) return; // columna no existe en el original -> queda en blanco
      const sourceCell = ws[XLSX.utils.encode_cell({ r, c: sourceCol })];
      if (!sourceCell) return;
      // copiamos la celda entera para preservar tipo/formato (horas, fechas, etc.)
      newWs[XLSX.utils.encode_cell({ r: outR, c: idx })] = { ...sourceCell };
    });
    rowCount++;
    maxRow = Math.max(maxRow, outR);
  }

  const lastCol = TARGET_COLUMNS.length - 1;
  newWs["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: lastCol } });

  // Ancho de columnas prolijo
  newWs["!cols"] = TARGET_COLUMNS.map((name) => ({ wch: Math.max(12, name.length + 2) }));

  return { sheet: newWs, rowCount, missingColumns };
}
