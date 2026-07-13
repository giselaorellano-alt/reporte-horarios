/**
 * CONFIGURACIÓN
 * ---------------------------------------------------------
 * Este es el ÚNICO archivo que necesitás tocar si el formato cambia.
 *
 * TARGET_COLUMNS: el orden y nombre exacto de las columnas que va a tener
 * el Excel de salida. Tiene que coincidir con el header (fila 2) del
 * archivo "Registros horarios 16-30 Junio", que es el formato que querés.
 *
 * Si Humand agrega o saca una columna del reporte, o si vos querés
 * cambiar el orden, editá este array. El script se adapta solo.
 *
 * Si una columna del array de acá abajo NO existe en el Excel que subís
 * (por ejemplo "Legajo", "Observaciones", "Horario ingreso programado",
 * "Horario egreso programado" no vienen en el reporte que descargás hoy),
 * el script la deja en blanco en la salida. Ni bien Humand la incluya,
 * se va a completar sola.
 */

const TARGET_COLUMNS = [
  "Usuario",
  "Nombre y apellido",
  "Legajo",
  "Sede de entrada",
  "Sede de salida",
  "Fecha",
  "Día",
  "Entrada",
  "Salida",
  "Licencias y/o permisos",
  "Observaciones",
  "Comentarios",
  "Fuera de ubicación",
  "Horas trabajadas por turno",
  "Horas trabajadas",
  "Horas programadas",
  "Balance",
  "Tiempo extra",
  "Feriados",
  "Ausencias injustificadas",
  "Incumplimientos",
  "Tardanza (hs)",
  "Horario ingreso programado",
  "Horario egreso programado",
  "Geolocalización entrada",
  "Geolocalización salida",
  "Reconocimiento facial omitido",
];

// Nombre que va a tener el archivo descargado (se le agrega la fecha/hora)
const OUTPUT_FILE_PREFIX = "Registros_horarios_reordenado";

// Texto del título que se pone en la fila 1 de la hoja de salida
const REPORT_TITLE = "Reporte detallado de Control horario";
