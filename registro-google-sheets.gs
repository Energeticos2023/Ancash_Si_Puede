/**
 * REGISTRO DE CAMPAÑA – BETTO 2026
 * Guarda en una hoja de Google los datos que llegan desde la página web:
 * Empadronamiento, Personeros y Sugerencias (una pestaña para cada uno).
 *
 * CÓMO INSTALARLO (5 minutos, desde una cuenta Google PERSONAL):
 * 1. Crea una hoja nueva en sheets.google.com y llámala "Registro Betto 2026".
 * 2. Menú Extensiones > Apps Script. Borra lo que aparece y pega TODO este código.
 * 3. Guarda (ícono de disquete) y ejecuta la función "prepararHojas" una vez.
 *    Google te pedirá permisos: acepta.
 * 4. Botón "Implementar" > "Nueva implementación" > tipo "Aplicación web".
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: Cualquier usuario
 *    Copia la URL que termina en /exec.
 * 5. Pega esa URL en la página web, en CONFIG.SCRIPT_URL.
 * Listo: cada registro aparecerá en tu hoja en tiempo real.
 */

const COLUMNAS = {
  'Empadronamiento': ['fecha','nombre','dni','edad','celular','whatsapp','provincia','distrito','centro_poblado','profesion','eje','intencion','voluntario','redes','reuniones','consentimiento'],
  'Personero':       ['fecha','nombre','dni','celular','whatsapp','provincia','distrito','local','mesa','experiencia','disponibilidad','capacitacion','consentimiento'],
  'Sugerencia':      ['fecha','provincia','distrito','sector','problema','propuesta','nombre','celular','consentimiento']
};

function prepararHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(COLUMNAS).forEach(nombre => {
    let h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
    if (h.getLastRow() === 0) {
      h.appendRow(COLUMNAS[nombre].map(c => c.toUpperCase()));
      h.getRange(1, 1, 1, COLUMNAS[nombre].length).setFontWeight('bold').setBackground('#0b2d6b').setFontColor('#ffffff');
      h.setFrozenRows(1);
    }
  });
  const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = e.parameter || {};
    const tipo = COLUMNAS[p.tipo] ? p.tipo : 'Sugerencia';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(tipo) || (prepararHojas(), ss.getSheetByName(tipo));
    // Evita fórmulas inyectadas: antepone un apóstrofo si el valor empieza con = + - @
    const limpio = v => { v = String(v || '').slice(0, 2000); return /^[=+\-@]/.test(v) ? "'" + v : v; };
    h.appendRow(COLUMNAS[tipo].map(c => limpio(p[c])));
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
