/**
 * Rifa de despensa: conecta la ruleta con esta hoja de cálculo.
 *
 * 1. Pega este código en Extensiones > Apps Script (borra lo que venga).
 * 2. Elige la función "configurar" arriba y dale "Ejecutar" una vez.
 *    Se crea la pestaña "Productos" con sus encabezados y la lista de abajo.
 *    Si después agregas productos a la lista PRODUCTOS, vuelve a ejecutar
 *    "configurar": solo agrega los que falten y no borra lo ya registrado.
 * 3. Implementar > Nueva implementación > Aplicación web
 *    (Ejecutar como: Yo; Quién tiene acceso: Cualquier persona).
 * 4. Copia la URL que termina en /exec y pégala en index.html.
 */

const HOJA = 'Productos';

const PRODUCTOS = [
  'Papel higiénico',
  'Detergente en polvo',
  'Suavizante de telas',
  'Cloro',
  'Leche',
  'Aceite',
  'Arroz',
  'Frijol',
  'Azúcar',
  'Harina de maíz',
  'Pasta para sopa',
  'Atún',
  'Puré de tomate',
  'Chiles en lata',
  'Elote en lata',
  'Frijoles en lata',
  'Consomé de pollo',
  'Mayonesa',
  'Sal',
  'Harina de trigo',
  'Servilletas',
  'Papel aluminio',
  'Plástico para envolver',
  'Cereal',
  'Galletas',
  'Café',
  'Chocolate en polvo o tablilla',
  'Avena',
  'Mermelada',
  'Jabón de tocador',
  'Shampoo',
  'Pasta de dientes',
  'Lavatrastes',
  'Jabón de barra para ropa',
  'Fibras',
  'Esponjas',
  'Bolsas para basura',
  'Toallas de cocina (servitoallas)',
  'Harina para hotcakes',
  'Trapos de microfibra'
];

function configurar() {
  const sh = hoja_();
  const ya = filas_(sh).map(function (x) { return x.producto.toLowerCase(); });
  const nuevos = PRODUCTOS.filter(function (p) { return ya.indexOf(p.toLowerCase()) < 0; });
  if (nuevos.length) {
    sh.getRange(sh.getLastRow() + 1, 1, nuevos.length, 1)
      .setValues(nuevos.map(function (p) { return [p]; }));
  }
}

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.getRange(1, 1, 1, 4)
      .setValues([['Producto', 'Nombre', 'Fecha', '¿Entregado?']])
      .setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 240);
    sh.setColumnWidth(2, 200);
  }
  return sh;
}

function filas_(sh) {
  const ultima = sh.getLastRow();
  if (ultima < 2) return [];
  return sh.getRange(2, 1, ultima - 1, 2).getValues()
    .map(function (r, i) {
      return { fila: i + 2, producto: String(r[0]).trim(), nombre: String(r[1]).trim() };
    })
    .filter(function (x) { return x.producto; });
}

function limpiar_(texto, max) {
  let t = String(texto || '').trim().slice(0, max);
  if (/^[=+\-@]/.test(t)) t = "'" + t; // evita que se interprete como fórmula
  return t;
}

function doGet() {
  const filas = filas_(hoja_());
  const libres = filas.filter(function (x) { return !x.nombre; }).map(function (x) { return x.producto; });
  return json_({ ok: true, libres: libres, total: filas.length });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const datos = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const nombre = limpiar_(datos.nombre, 60);
    if (!nombre) return json_({ ok: false, error: 'falta_nombre' });

    const sh = hoja_();
    const filas = filas_(sh);
    const libres = filas.filter(function (x) { return !x.nombre; });
    if (!libres.length) return json_({ ok: false, error: 'sin_productos', libres: [], total: filas.length });

    const elegido = libres[Math.floor(Math.random() * libres.length)];
    sh.getRange(elegido.fila, 2, 1, 2).setValues([[nombre, new Date()]]);
    SpreadsheetApp.flush();

    const restantes = libres
      .filter(function (x) { return x.fila !== elegido.fila; })
      .map(function (x) { return x.producto; });
    return json_({ ok: true, producto: elegido.producto, libres: restantes, total: filas.length });
  } catch (err) {
    return json_({ ok: false, error: 'error_servidor' });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
