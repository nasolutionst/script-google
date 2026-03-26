const CONFIG = {
  HOJA_FACTURA: 'Factura',
  HOJA_PRESUPUESTO: 'Presupuesto',
  HOJA_BASE: 'BaseDropdown',
  HOJA_REGISTRO: 'RegistroFacturas',
  HOJA_REGISTRO_PRESUPUESTOS: 'RegistroPresupuestos',
  
  PDF_FOLDER_ID: '1qDlTzogwfnQlnycc83wmsQIC2apf4ZzO',
  NOMBRE_CARPETA_FACTURAS_CLIENTE: 'Facturas por cliente',
  NOMBRE_CARPETA_PRESUPUESTOS_CLIENTE: 'Presupuestos por cliente',

  // Puede ser el ID real de la carpeta o simplemente el nombre de la carpeta raíz
  PDF_FOLDER_MENSUAL_ID: 'Facturas por mes',

  RANGO_NUMERO_FACTURA: 'E2',
  RANGO_FECHA_FACTURA: 'E3',
  RANGO_FECHA_VENCIMIENTO: 'E4',
  RANGO_CLIENTE: 'C20',
  RANGO_CIF: 'C21',
  RANGO_DIRECCION: 'C22',

  FILA_PRIMERA_LINEA: 25,
  FILA_ULTIMA_LINEA_BASE: 29,
  TEXTO_TOTALES: 'Totales:',

  IVA_POR_DEFECTO: 0.21,
  PLAZO_VENCIMIENTO_DIAS: 21,
  PLAZO_VALIDEZ_PRESUPUESTO_DIAS: 30,

  ESTADO_FACTURA: 'Emitida',
  ESTADO_PRESUPUESTO: 'Pendiente',

  PREFIJO_PRESUPUESTO: 'PRE',
  FORMATO_FECHA: 'dd/MM/yyyy'
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Facturas')
    .addItem('Guardar factura en registro', 'guardarFacturaEnRegistro')
    .addItem('Nueva factura', 'nuevaFactura')
    .addItem('Añadir línea de producto/servicio', 'anadirLineaProductoServicio')
    .addItem('Ir a factura guardada', 'irAFacturaGuardada')
    .addItem('Exportar factura a PDF', 'exportarFacturaPDF')
    .addSeparator()
    .addItem('Reparar factura actual', 'repararFacturaActual')
    .addToUi();

  ui.createMenu('Presupuestos')
    .addItem('Guardar presupuesto en registro', 'guardarPresupuestoEnRegistro')
    .addItem('Nuevo presupuesto', 'nuevoPresupuesto')
    .addItem('Añadir línea de producto/servicio', 'anadirLineaProductoServicio')
    .addItem('Ir a presupuesto guardado', 'irAPresupuestoGuardado')
    .addItem('Exportar presupuesto a PDF', 'exportarPresupuestoPDF')
    .addSeparator()
    .addItem('Convertir presupuesto en factura', 'convertirPresupuestoAFactura')
    .addItem('Reparar presupuesto actual', 'repararPresupuestoActual')
    .addToUi();

  const plantillaFactura = obtenerHojaPlantillaFactura_();
  if (plantillaFactura) {
    repararHojaFactura_(plantillaFactura);
  }

  const plantillaPresupuesto = obtenerHojaPlantillaPresupuesto_();
  if (plantillaPresupuesto) {
    repararHojaFactura_(plantillaPresupuesto);
  }

  repararDropdownClientes_();
}

function onEdit(e) {
  if (!e || !e.range) return;

  const hoja = e.range.getSheet();
  const fila = e.range.getRow();
  const columna = e.range.getColumn();
  const a1 = e.range.getA1Notation();

  // Si se edita la base de clientes, reconstruye el desplegable al momento
  if (hoja.getName() === CONFIG.HOJA_BASE) {
    if (fila >= 2 && columna === 1) {
      repararDropdownClientes_();
    }
    return;
  }

  if (!esHojaDocumentoEditable_(hoja)) return;

  const layout = obtenerLayoutFactura_(hoja);

  if (a1 === CONFIG.RANGO_CLIENTE) {
    rellenarDatosCliente_(hoja);
  }

  if (a1 === CONFIG.RANGO_FECHA_FACTURA) {
    actualizarFechaSecundariaDocumento_(hoja);
  }

  const editandoLineas =
    fila >= CONFIG.FILA_PRIMERA_LINEA &&
    fila <= layout.filaUltimaLinea &&
    columna >= 1 &&
    columna <= 6;

  const editandoDescuento =
    fila === layout.filaResumenValores && columna === 1;

  const editandoIvaResumen =
    fila === layout.filaResumenValores && columna === 4;

  if (
    editandoLineas ||
    editandoDescuento ||
    editandoIvaResumen ||
    a1 === CONFIG.RANGO_CLIENTE ||
    a1 === CONFIG.RANGO_FECHA_FACTURA
  ) {
    recalcularFactura_(hoja);
  }
}

/* =========================
   FACTURAS
========================= */

function guardarFacturaEnRegistro() {
  const hojaFactura = obtenerHojaFacturaActiva_();
  guardarFacturaEnRegistroInterno_(hojaFactura, true);
}

function guardarFacturaEnRegistroInterno_(hojaFactura, mostrarAlerta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistro = ss.getSheetByName(CONFIG.HOJA_REGISTRO);
  const ui = SpreadsheetApp.getUi();

  if (!hojaFactura) {
    ui.alert('No hay una hoja de factura activa.');
    return null;
  }

  if (!hojaRegistro) {
    ui.alert(`No existe la hoja "${CONFIG.HOJA_REGISTRO}".`);
    return null;
  }

  repararHojaFactura_(hojaFactura);
  SpreadsheetApp.flush();

  const layout = obtenerLayoutFactura_(hojaFactura);

  const numeroFactura = hojaFactura.getRange(CONFIG.RANGO_NUMERO_FACTURA).getDisplayValue().trim();
  const cliente = hojaFactura.getRange(CONFIG.RANGO_CLIENTE).getDisplayValue().trim();
  const cif = hojaFactura.getRange(CONFIG.RANGO_CIF).getDisplayValue().trim();
  const direccion = hojaFactura.getRange(CONFIG.RANGO_DIRECCION).getDisplayValue().trim();
  const fechaFactura = hojaFactura.getRange(CONFIG.RANGO_FECHA_FACTURA).getValue();

  const baseImponible = hojaFactura.getRange(`C${layout.filaResumenValores}`).getValue();
  const iva = hojaFactura.getRange(`E${layout.filaResumenValores}`).getValue();
  const total = hojaFactura.getRange(`F${layout.filaResumenValores}`).getValue();

  if (!numeroFactura) {
    ui.alert(`Falta el número de factura en ${CONFIG.RANGO_NUMERO_FACTURA}.`);
    return null;
  }

  if (!cliente) {
    ui.alert(`Falta seleccionar el cliente en ${CONFIG.RANGO_CLIENTE}.`);
    return null;
  }

  if (!(fechaFactura instanceof Date) || isNaN(fechaFactura.getTime())) {
    ui.alert(`La fecha de emisión en ${CONFIG.RANGO_FECHA_FACTURA} no es válida.`);
    return null;
  }

  if (!hayConceptosFactura_(hojaFactura)) {
    ui.alert('No hay ningún concepto en la factura. Rellena al menos una línea antes de guardarla.');
    return null;
  }

  const datosFila = [
    new Date(),
    numeroFactura,
    fechaFactura,
    cliente,
    cif,
    direccion,
    baseImponible,
    iva,
    total,
    CONFIG.ESTADO_FACTURA
  ];

  const filaExistente = buscarFilaDocumentoEnRegistro_(hojaRegistro, numeroFactura);

  let accionRegistro = '';
  if (filaExistente) {
    hojaRegistro.getRange(filaExistente, 1, 1, datosFila.length).setValues([datosFila]);
    accionRegistro = 'actualizada';
  } else {
    hojaRegistro.appendRow(datosFila);
    accionRegistro = 'guardada';
  }

  let hojaArchivo = null;
  let mensajeExtra = '';

  if (hojaFactura.getName() === CONFIG.HOJA_FACTURA) {
    hojaArchivo = crearOActualizarHojaArchivo_(hojaFactura, numeroFactura);
    mensajeExtra = `\nSe ha creado/actualizado la pestaña editable: ${hojaArchivo.getName()}`;
  }

  if (mostrarAlerta) {
    ui.alert(`Factura ${numeroFactura} ${accionRegistro} correctamente en RegistroFacturas.${mensajeExtra}`);
  }

  return {
    numeroFactura,
    accionRegistro,
    hojaArchivo
  };
}

function nuevaFactura() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaFactura = obtenerHojaPlantillaFactura_();
  const ui = SpreadsheetApp.getUi();

  if (!hojaFactura) {
    ui.alert(`No existe la hoja "${CONFIG.HOJA_FACTURA}".`);
    return;
  }

  const respuesta = ui.alert(
    'Nueva factura',
    'Se limpiarán los datos editables de la plantilla y se volverá a empezar con una sola línea. ¿Quieres continuar?',
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) {
    return;
  }

  ss.setActiveSheet(hojaFactura);
  prepararDocumentoNuevo_(hojaFactura, 'factura');

  ui.alert('La plantilla "Factura" ha quedado lista para una nueva factura.');
}

function exportarFacturaPDF() {
  const hoja = obtenerHojaFacturaActiva_();
  exportarDocumentoPDF_(hoja, CONFIG.NOMBRE_CARPETA_FACTURAS_CLIENTE);
}

function irAFacturaGuardada() {
  irADocumentoGuardado_(
    'Ir a factura guardada',
    'Escribe el número exacto de factura, por ejemplo: FAC-2026-001'
  );
}

function repararFacturaActual() {
  const hojaFactura = obtenerHojaFacturaActiva_();
  if (!hojaFactura) return;
  repararHojaFactura_(hojaFactura);
}

/* =========================
   PRESUPUESTOS
========================= */

function guardarPresupuestoEnRegistro() {
  const hojaPresupuesto = obtenerHojaPresupuestoActiva_();
  guardarPresupuestoEnRegistroInterno_(hojaPresupuesto, true);
}

function guardarPresupuestoEnRegistroInterno_(hojaPresupuesto, mostrarAlerta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistro = ss.getSheetByName(CONFIG.HOJA_REGISTRO_PRESUPUESTOS);
  const ui = SpreadsheetApp.getUi();

  if (!hojaPresupuesto) {
    ui.alert('No hay una hoja de presupuesto activa.');
    return null;
  }

  if (!hojaRegistro) {
    ui.alert(`No existe la hoja "${CONFIG.HOJA_REGISTRO_PRESUPUESTOS}".`);
    return null;
  }

  repararHojaFactura_(hojaPresupuesto);
  SpreadsheetApp.flush();

  const layout = obtenerLayoutFactura_(hojaPresupuesto);

  let numeroPresupuesto = hojaPresupuesto.getRange(CONFIG.RANGO_NUMERO_FACTURA).getDisplayValue().trim();
  if (!numeroPresupuesto) {
    numeroPresupuesto = asignarNumeroPresupuestoSiFalta_(hojaPresupuesto);
    SpreadsheetApp.flush();
  }

  const cliente = hojaPresupuesto.getRange(CONFIG.RANGO_CLIENTE).getDisplayValue().trim();
  const cif = hojaPresupuesto.getRange(CONFIG.RANGO_CIF).getDisplayValue().trim();
  const direccion = hojaPresupuesto.getRange(CONFIG.RANGO_DIRECCION).getDisplayValue().trim();
  const fechaPresupuesto = hojaPresupuesto.getRange(CONFIG.RANGO_FECHA_FACTURA).getValue();
  const fechaValidez = hojaPresupuesto.getRange(CONFIG.RANGO_FECHA_VENCIMIENTO).getValue();

  const baseImponible = hojaPresupuesto.getRange(`C${layout.filaResumenValores}`).getValue();
  const iva = hojaPresupuesto.getRange(`E${layout.filaResumenValores}`).getValue();
  const total = hojaPresupuesto.getRange(`F${layout.filaResumenValores}`).getValue();

  if (!numeroPresupuesto) {
    ui.alert(`Falta el número de presupuesto en ${CONFIG.RANGO_NUMERO_FACTURA}.`);
    return null;
  }

  if (!cliente) {
    ui.alert(`Falta seleccionar el cliente en ${CONFIG.RANGO_CLIENTE}.`);
    return null;
  }

  if (!(fechaPresupuesto instanceof Date) || isNaN(fechaPresupuesto.getTime())) {
    ui.alert(`La fecha del presupuesto en ${CONFIG.RANGO_FECHA_FACTURA} no es válida.`);
    return null;
  }

  if (!hayConceptosFactura_(hojaPresupuesto)) {
    ui.alert('No hay ningún concepto en el presupuesto. Rellena al menos una línea antes de guardarlo.');
    return null;
  }

  const filaExistente = buscarFilaDocumentoEnRegistro_(hojaRegistro, numeroPresupuesto);

  let estado = CONFIG.ESTADO_PRESUPUESTO;
  let facturaVinculada = '';
  let fechaConversion = '';

  if (filaExistente) {
    estado = hojaRegistro.getRange(filaExistente, 11).getDisplayValue().trim() || CONFIG.ESTADO_PRESUPUESTO;
    facturaVinculada = hojaRegistro.getRange(filaExistente, 12).getDisplayValue().trim();
    fechaConversion = hojaRegistro.getRange(filaExistente, 13).getValue();
  }

  const datosFila = [
    new Date(),
    numeroPresupuesto,
    fechaPresupuesto,
    fechaValidez,
    cliente,
    cif,
    direccion,
    baseImponible,
    iva,
    total,
    estado,
    facturaVinculada,
    fechaConversion || ''
  ];

  let accionRegistro = '';
  if (filaExistente) {
    hojaRegistro.getRange(filaExistente, 1, 1, datosFila.length).setValues([datosFila]);
    accionRegistro = 'actualizado';
  } else {
    hojaRegistro.appendRow(datosFila);
    accionRegistro = 'guardado';
  }

  let hojaArchivo = null;
  let mensajeExtra = '';

  if (hojaPresupuesto.getName() === CONFIG.HOJA_PRESUPUESTO) {
    hojaArchivo = crearOActualizarHojaArchivo_(hojaPresupuesto, numeroPresupuesto);
    mensajeExtra = `\nSe ha creado/actualizado la pestaña editable: ${hojaArchivo.getName()}`;
  }

  if (mostrarAlerta) {
    ui.alert(`Presupuesto ${numeroPresupuesto} ${accionRegistro} correctamente en RegistroPresupuestos.${mensajeExtra}`);
  }

  return {
    numeroPresupuesto,
    accionRegistro,
    hojaArchivo
  };
}

function nuevoPresupuesto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPresupuesto = obtenerHojaPlantillaPresupuesto_();
  const ui = SpreadsheetApp.getUi();

  if (!hojaPresupuesto) {
    ui.alert(`No existe la hoja "${CONFIG.HOJA_PRESUPUESTO}".`);
    return;
  }

  const respuesta = ui.alert(
    'Nuevo presupuesto',
    'Se limpiarán los datos editables de la plantilla y se generará un nuevo número de presupuesto. ¿Quieres continuar?',
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) {
    return;
  }

  ss.setActiveSheet(hojaPresupuesto);
  prepararDocumentoNuevo_(hojaPresupuesto, 'presupuesto');

  ui.alert('La plantilla "Presupuesto" ha quedado lista para un nuevo presupuesto.');
}

function exportarPresupuestoPDF() {
  const hoja = obtenerHojaPresupuestoActiva_();
  exportarDocumentoPDF_(hoja, CONFIG.NOMBRE_CARPETA_PRESUPUESTOS_CLIENTE);
}

function irAPresupuestoGuardado() {
  irADocumentoGuardado_(
    'Ir a presupuesto guardado',
    'Escribe el número exacto de presupuesto, por ejemplo: PRE-2026-001'
  );
}

function repararPresupuestoActual() {
  const hojaPresupuesto = obtenerHojaPresupuestoActiva_();
  if (!hojaPresupuesto) return;
  repararHojaFactura_(hojaPresupuesto);
}

function convertirPresupuestoAFactura() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const hojaPresupuesto = obtenerHojaPresupuestoActiva_();
  const hojaFactura = obtenerHojaPlantillaFactura_();
  const hojaRegistroPres = ss.getSheetByName(CONFIG.HOJA_REGISTRO_PRESUPUESTOS);

  if (!hojaPresupuesto) {
    ui.alert('No hay una hoja de presupuesto activa.');
    return;
  }

  if (!hojaFactura) {
    ui.alert(`No existe la hoja "${CONFIG.HOJA_FACTURA}".`);
    return;
  }

  if (!hojaRegistroPres) {
    ui.alert(`No existe la hoja "${CONFIG.HOJA_REGISTRO_PRESUPUESTOS}".`);
    return;
  }

  const guardadoPres = guardarPresupuestoEnRegistroInterno_(hojaPresupuesto, false);
  if (!guardadoPres) return;

  const numeroPresupuesto = guardadoPres.numeroPresupuesto;
  const filaPresupuesto = buscarFilaDocumentoEnRegistro_(hojaRegistroPres, numeroPresupuesto);

  if (filaPresupuesto) {
    const facturaYaVinculada = hojaRegistroPres.getRange(filaPresupuesto, 12).getDisplayValue().trim();
    if (facturaYaVinculada) {
      ui.alert(`Este presupuesto ya está convertido en la factura ${facturaYaVinculada}.`);
      return;
    }
  }

  const respuesta = ui.alert(
    'Convertir presupuesto en factura',
    `Se copiarán los datos del presupuesto ${numeroPresupuesto} a la plantilla de factura y se guardará la factura en el registro. ¿Quieres continuar?`,
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) {
    return;
  }

  copiarPresupuestoAFactura_(hojaPresupuesto, hojaFactura);
  ss.setActiveSheet(hojaFactura);
  SpreadsheetApp.flush();

  const guardadoFactura = guardarFacturaEnRegistroInterno_(hojaFactura, false);
  if (!guardadoFactura) {
    ui.alert('El presupuesto se copió a la factura, pero no se pudo guardar la factura en el registro.');
    return;
  }

  actualizarRegistroPresupuestoFacturado_(numeroPresupuesto, guardadoFactura.numeroFactura);

  ui.alert(
    `Presupuesto ${numeroPresupuesto} convertido correctamente en factura ${guardadoFactura.numeroFactura}.`
  );
}

/* =========================
   FUNCIONES COMPARTIDAS
========================= */

function prepararDocumentoNuevo_(hoja, tipoDocumento) {
  restablecerLineasFactura_(hoja);

  const layout = obtenerLayoutFactura_(hoja);

  hoja.getRange('C20:C22').clearContent();
  hoja.getRange(`B${layout.filaNotas}`).clearContent();

  hoja.getRange(`A${layout.filaResumenValores}`).setValue(0);
  hoja.getRange(`D${layout.filaResumenValores}`).setValue(CONFIG.IVA_POR_DEFECTO);

  hoja.getRange(CONFIG.RANGO_FECHA_FACTURA).setValue(new Date());

  if (tipoDocumento === 'presupuesto') {
    forzarNuevoNumeroPresupuesto_(hoja);
  }

  actualizarFechaSecundariaDocumento_(hoja);
  recalcularFactura_(hoja);
  aplicarFormatosFactura_(hoja);

  hoja.setActiveRange(hoja.getRange(`A${CONFIG.FILA_PRIMERA_LINEA}`));
}

function anadirLineaProductoServicio() {
  const hojaDocumento = obtenerHojaDocumentoActiva_();
  const ui = SpreadsheetApp.getUi();

  if (!hojaDocumento) {
    ui.alert('Activa primero una hoja de factura o de presupuesto.');
    return;
  }

  let layout = obtenerLayoutFactura_(hojaDocumento);

  for (let fila = CONFIG.FILA_PRIMERA_LINEA + 1; fila <= Math.min(CONFIG.FILA_ULTIMA_LINEA_BASE, layout.filaUltimaLinea); fila++) {
    if (hojaDocumento.isRowHiddenByUser(fila)) {
      hojaDocumento.showRows(fila, 1);
      hojaDocumento.getRange(`A${fila}:F${fila}`).clearContent();
      hojaDocumento.getRange(`E${fila}`).setValue(CONFIG.IVA_POR_DEFECTO);
      aplicarFormatosFactura_(hojaDocumento);
      recalcularFactura_(hojaDocumento);
      hojaDocumento.setActiveRange(hojaDocumento.getRange(`A${fila}`));
      return;
    }
  }

  insertarLineaDocumento_(hojaDocumento, true);
}

function insertarLineaDocumento_(hojaDocumento, recalcularDespues) {
  const layout = obtenerLayoutFactura_(hojaDocumento);
  const filaInsertar = layout.filaTotalesTitulo;
  const filaModelo = filaInsertar - 1;

  hojaDocumento.insertRowBefore(filaInsertar);
  hojaDocumento.getRange(`A${filaModelo}:F${filaModelo}`).copyFormatToRange(hojaDocumento, 1, 6, filaInsertar, filaInsertar);
  hojaDocumento.setRowHeight(filaInsertar, hojaDocumento.getRowHeight(filaModelo));
  hojaDocumento.getRange(`A${filaInsertar}:F${filaInsertar}`).clearContent();
  hojaDocumento.getRange(`E${filaInsertar}`).setValue(CONFIG.IVA_POR_DEFECTO);

  aplicarFormatosFactura_(hojaDocumento);

  if (recalcularDespues) {
    recalcularFactura_(hojaDocumento);
    hojaDocumento.setActiveRange(hojaDocumento.getRange(`A${filaInsertar}`));
  }
}

function exportarDocumentoPDF_(hoja, nombreCarpetaRaizClientes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (!hoja) {
    ui.alert('No hay una hoja activa de factura o presupuesto.');
    return;
  }

  repararHojaFactura_(hoja);
  SpreadsheetApp.flush();

  const numeroDocumento = hoja.getRange(CONFIG.RANGO_NUMERO_FACTURA).getDisplayValue().trim();
  const cliente = hoja.getRange(CONFIG.RANGO_CLIENTE).getDisplayValue().trim();
  const fechaDocumento = hoja.getRange(CONFIG.RANGO_FECHA_FACTURA).getValue();
  const tipoDocumento = obtenerTipoDocumentoDeHoja_(hoja);

  if (!numeroDocumento) {
    ui.alert('Falta el número del documento.');
    return;
  }

  if (!cliente) {
    ui.alert('Falta el nombre del cliente.');
    return;
  }

  if (!hayConceptosFactura_(hoja)) {
    ui.alert('No hay conceptos en el documento. No se puede exportar un PDF vacío.');
    return;
  }

  if (
    tipoDocumento === 'factura' &&
    (!(fechaDocumento instanceof Date) || isNaN(fechaDocumento.getTime()))
  ) {
    ui.alert('La fecha de la factura no es válida. No se puede clasificar por mes.');
    return;
  }

  let carpetaGeneral;
  try {
    carpetaGeneral = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);
  } catch (error) {
    ui.alert('No se pudo acceder a la carpeta general de PDFs configurada.');
    return;
  }

  const carpetaRaizClientes = obtenerOCrearCarpetaRaizPorNombre_(nombreCarpetaRaizClientes);
  if (!carpetaRaizClientes) {
    ui.alert(`No se pudo localizar ni crear la carpeta raíz "${nombreCarpetaRaizClientes}".`);
    return;
  }

  const carpetaCliente = obtenerOCrearSubcarpeta_(carpetaRaizClientes, cliente);

  const spreadsheetId = ss.getId();
  const sheetId = hoja.getSheetId();
  const nombreArchivo = `${sanitizarNombreArchivo_(numeroDocumento)} - ${sanitizarNombreArchivo_(cliente)}.pdf`;

  const url =
    `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export` +
    `?format=pdf` +
    `&gid=${sheetId}` +
    `&size=A4` +
    `&portrait=true` +
    `&fitw=true` +
    `&sheetnames=false` +
    `&printtitle=false` +
    `&pagenumbers=false` +
    `&gridlines=false` +
    `&fzr=false` +
    `&top_margin=0.50` +
    `&bottom_margin=0.50` +
    `&left_margin=0.50` +
    `&right_margin=0.50`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    ui.alert('No se pudo generar el PDF. Revisa permisos y configuración de las carpetas.');
    return;
  }

  const blob = response.getBlob().setName(nombreArchivo);

  // Carpeta general
  guardarOReemplazarArchivoEnCarpeta_(carpetaGeneral, blob, nombreArchivo);

  // Carpeta por cliente
  guardarOReemplazarArchivoEnCarpeta_(carpetaCliente, blob, nombreArchivo);

  // Carpeta por mes solo para facturas
  let mensajeMes = '';
  if (tipoDocumento === 'factura') {
    const carpetaRaizMeses = obtenerCarpetaRaizMensualFacturas_();
    const nombreMes = obtenerNombreMes_(fechaDocumento);
    const carpetaMes = obtenerOCrearSubcarpeta_(carpetaRaizMeses, nombreMes);

    // Borra versiones antiguas del mismo PDF en cualquier mes
    borrarArchivoPorNombreEnArbol_(carpetaRaizMeses, nombreArchivo);

    // Guarda la nueva versión en el mes correcto
    guardarOReemplazarArchivoEnCarpeta_(carpetaMes, blob, nombreArchivo);

    mensajeMes = `\n- ${obtenerNombreVisibleCarpetaMensual_()}/${nombreMes}`;
  }

  ui.alert(
    `PDF guardado correctamente:\n` +
    `- Carpeta general\n` +
    `- ${nombreCarpetaRaizClientes}/${cliente}` +
    `${mensajeMes}`
  );
}

function irADocumentoGuardado_(titulo, mensajePrompt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const respuesta = ui.prompt(
    titulo,
    mensajePrompt,
    ui.ButtonSet.OK_CANCEL
  );

  if (respuesta.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const numeroDocumento = respuesta.getResponseText().trim();
  if (!numeroDocumento) {
    ui.alert('No has escrito ningún número.');
    return;
  }

  const hoja = ss.getSheetByName(normalizarNombreHojaDocumento_(numeroDocumento));
  if (!hoja) {
    ui.alert(`No existe una pestaña archivada para ${numeroDocumento}.`);
    return;
  }

  ss.setActiveSheet(hoja);
  hoja.setActiveRange(hoja.getRange(`A${CONFIG.FILA_PRIMERA_LINEA}`));
}

function copiarPresupuestoAFactura_(hojaOrigen, hojaDestino) {
  repararHojaFactura_(hojaOrigen);

  const layouts = replicarLayoutLineasDocumento_(hojaOrigen, hojaDestino);
  const layoutOrigen = layouts.layoutOrigen;
  const layoutDestino = layouts.layoutDestino;

  const lineasOrigen = hojaOrigen.getRange(`A${CONFIG.FILA_PRIMERA_LINEA}:F${layoutOrigen.filaUltimaLinea}`).getValues();
  hojaDestino.getRange(`A${CONFIG.FILA_PRIMERA_LINEA}:F${layoutDestino.filaUltimaLinea}`).setValues(lineasOrigen);

  hojaDestino.getRange(CONFIG.RANGO_CLIENTE).setValue(
    hojaOrigen.getRange(CONFIG.RANGO_CLIENTE).getDisplayValue()
  );
  hojaDestino.getRange(CONFIG.RANGO_CIF).setValue(
    hojaOrigen.getRange(CONFIG.RANGO_CIF).getDisplayValue()
  );
  hojaDestino.getRange(CONFIG.RANGO_DIRECCION).setValue(
    hojaOrigen.getRange(CONFIG.RANGO_DIRECCION).getDisplayValue()
  );

  hojaDestino.getRange(CONFIG.RANGO_FECHA_FACTURA).setValue(new Date());

  hojaDestino.getRange(`A${layoutDestino.filaResumenValores}`).setValue(
    hojaOrigen.getRange(`A${layoutOrigen.filaResumenValores}`).getValue()
  );
  hojaDestino.getRange(`D${layoutDestino.filaResumenValores}`).setValue(
    hojaOrigen.getRange(`D${layoutOrigen.filaResumenValores}`).getValue()
  );

  hojaDestino.getRange(`B${layoutDestino.filaNotas}`).setValue(
    hojaOrigen.getRange(`B${layoutOrigen.filaNotas}`).getValue()
  );

  actualizarFechaSecundariaDocumento_(hojaDestino);
  recalcularFactura_(hojaDestino);
  aplicarFormatosFactura_(hojaDestino);
}

function replicarLayoutLineasDocumento_(hojaOrigen, hojaDestino) {
  const layoutOrigen = obtenerLayoutFactura_(hojaOrigen);

  restablecerLineasFactura_(hojaDestino);

  const totalLineasOrigen = layoutOrigen.filaUltimaLinea - CONFIG.FILA_PRIMERA_LINEA + 1;
  const totalLineasBase = CONFIG.FILA_ULTIMA_LINEA_BASE - CONFIG.FILA_PRIMERA_LINEA + 1;
  const lineasExtra = Math.max(0, totalLineasOrigen - totalLineasBase);

  for (let i = 0; i < lineasExtra; i++) {
    insertarLineaDocumento_(hojaDestino, false);
  }

  const layoutDestino = obtenerLayoutFactura_(hojaDestino);

  hojaDestino.showRows(
    CONFIG.FILA_PRIMERA_LINEA,
    layoutDestino.filaUltimaLinea - CONFIG.FILA_PRIMERA_LINEA + 1
  );

  for (let fila = CONFIG.FILA_PRIMERA_LINEA; fila <= layoutOrigen.filaUltimaLinea; fila++) {
    if (hojaOrigen.isRowHiddenByUser(fila)) {
      hojaDestino.hideRows(fila, 1);
    }
  }

  return { layoutOrigen, layoutDestino };
}

function actualizarRegistroPresupuestoFacturado_(numeroPresupuesto, numeroFactura) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistro = ss.getSheetByName(CONFIG.HOJA_REGISTRO_PRESUPUESTOS);
  if (!hojaRegistro) return;

  const fila = buscarFilaDocumentoEnRegistro_(hojaRegistro, numeroPresupuesto);
  if (!fila) return;

  hojaRegistro.getRange(fila, 11).setValue('Facturado');
  hojaRegistro.getRange(fila, 12).setValue(numeroFactura);
  hojaRegistro.getRange(fila, 13).setValue(new Date());
}

/* =========================
   DATOS CLIENTE / FECHAS / CÁLCULOS
========================= */

function rellenarDatosCliente_(hojaFactura) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaBase = ss.getSheetByName(CONFIG.HOJA_BASE);

  if (!hojaBase) return;

  const cliente = hojaFactura.getRange(CONFIG.RANGO_CLIENTE).getDisplayValue().trim();

  if (!cliente) {
    hojaFactura.getRange(CONFIG.RANGO_CIF).clearContent();
    hojaFactura.getRange(CONFIG.RANGO_DIRECCION).clearContent();
    return;
  }

  const ultimaFila = hojaBase.getLastRow();
  if (ultimaFila < 2) {
    hojaFactura.getRange(CONFIG.RANGO_CIF).clearContent();
    hojaFactura.getRange(CONFIG.RANGO_DIRECCION).clearContent();
    return;
  }

  const datos = hojaBase.getRange(2, 1, ultimaFila - 1, 4).getValues();
  const filaCliente = datos.find(fila => String(fila[0]).trim() === cliente);

  hojaFactura.getRange(CONFIG.RANGO_CIF).setValue(filaCliente ? filaCliente[1] : '');
  hojaFactura.getRange(CONFIG.RANGO_DIRECCION).setValue(filaCliente ? filaCliente[3] : '');
}

function actualizarFechaSecundariaDocumento_(hojaFactura) {
  const layout = obtenerLayoutFactura_(hojaFactura);
  const fechaEmision = hojaFactura.getRange(CONFIG.RANGO_FECHA_FACTURA).getValue();

  if (!(fechaEmision instanceof Date) || isNaN(fechaEmision.getTime())) {
    hojaFactura.getRange(CONFIG.RANGO_FECHA_VENCIMIENTO).clearContent();
    hojaFactura.getRange(`E${layout.filaDistribucionValores}`).clearContent();
    return;
  }

  const tipoDocumento = obtenerTipoDocumentoDeHoja_(hojaFactura);
  const dias = tipoDocumento === 'presupuesto'
    ? CONFIG.PLAZO_VALIDEZ_PRESUPUESTO_DIAS
    : CONFIG.PLAZO_VENCIMIENTO_DIAS;

  const fechaSecundaria = new Date(fechaEmision);
  fechaSecundaria.setDate(fechaSecundaria.getDate() + dias);

  hojaFactura.getRange(CONFIG.RANGO_FECHA_VENCIMIENTO).setValue(fechaSecundaria);
  hojaFactura.getRange(`E${layout.filaDistribucionValores}`).setValue(fechaSecundaria);
}

function recalcularFactura_(hojaFactura) {
  const layout = obtenerLayoutFactura_(hojaFactura);
  const lineas = hojaFactura.getRange(`A${CONFIG.FILA_PRIMERA_LINEA}:F${layout.filaUltimaLinea}`).getValues();

  const importes = [];
  const ivas = [];
  const totalesLinea = [];

  let neto = 0;

  for (let i = 0; i < lineas.length; i++) {
    const filaReal = CONFIG.FILA_PRIMERA_LINEA + i;

    if (hojaFactura.isRowHiddenByUser(filaReal)) {
      importes.push(['']);
      ivas.push([CONFIG.IVA_POR_DEFECTO]);
      totalesLinea.push(['']);
      continue;
    }

    const descripcion = String(lineas[i][0] || '').trim();
    const cantidad = normalizarNumero_(lineas[i][1]);
    const precio = normalizarNumero_(lineas[i][2]);
    let ivaLinea = normalizarPorcentaje_(lineas[i][4]);

    const filaTieneAlgo =
      descripcion !== '' ||
      (lineas[i][1] !== '' && lineas[i][1] !== null) ||
      (lineas[i][2] !== '' && lineas[i][2] !== null);

    if (ivaLinea === null) {
      ivaLinea = CONFIG.IVA_POR_DEFECTO;
    }

    if (!filaTieneAlgo) {
      importes.push(['']);
      ivas.push([ivaLinea]);
      totalesLinea.push(['']);
      continue;
    }

    if (cantidad === null || precio === null) {
      importes.push(['']);
      ivas.push([ivaLinea]);
      totalesLinea.push(['']);
      continue;
    }

    const importe = cantidad * precio;
    const totalLinea = importe * (1 + ivaLinea);

    neto += importe;

    importes.push([importe]);
    ivas.push([ivaLinea]);
    totalesLinea.push([totalLinea]);
  }

  hojaFactura.getRange(`D${CONFIG.FILA_PRIMERA_LINEA}:D${layout.filaUltimaLinea}`).setValues(importes);
  hojaFactura.getRange(`E${CONFIG.FILA_PRIMERA_LINEA}:E${layout.filaUltimaLinea}`).setValues(ivas);
  hojaFactura.getRange(`F${CONFIG.FILA_PRIMERA_LINEA}:F${layout.filaUltimaLinea}`).setValues(totalesLinea);

  let descuento = normalizarPorcentaje_(hojaFactura.getRange(`A${layout.filaResumenValores}`).getValue());
  if (descuento === null) descuento = 0;

  let ivaResumen = normalizarPorcentaje_(hojaFactura.getRange(`D${layout.filaResumenValores}`).getValue());
  if (ivaResumen === null) ivaResumen = CONFIG.IVA_POR_DEFECTO;

  const baseImponible = neto * (1 - descuento);
  const importeIva = baseImponible * ivaResumen;
  const total = baseImponible + importeIva;

  hojaFactura.getRange(`A${layout.filaResumenValores}`).setValue(descuento);
  hojaFactura.getRange(`D${layout.filaResumenValores}`).setValue(ivaResumen);

  hojaFactura.getRange(`B${layout.filaResumenValores}`).setValue(neto);
  hojaFactura.getRange(`C${layout.filaResumenValores}`).setValue(baseImponible);
  hojaFactura.getRange(`E${layout.filaResumenValores}`).setValue(importeIva);
  hojaFactura.getRange(`F${layout.filaResumenValores}`).setValue(total);
  hojaFactura.getRange(`F${layout.filaDistribucionValores}`).setValue(total);

  aplicarFormatosFactura_(hojaFactura);
}

function aplicarFormatosFactura_(hojaFactura) {
  const layout = obtenerLayoutFactura_(hojaFactura);

  hojaFactura.getRange(CONFIG.RANGO_FECHA_FACTURA).setNumberFormat(CONFIG.FORMATO_FECHA);
  hojaFactura.getRange(CONFIG.RANGO_FECHA_VENCIMIENTO).setNumberFormat(CONFIG.FORMATO_FECHA);
  hojaFactura.getRange(`E${layout.filaDistribucionValores}`).setNumberFormat(CONFIG.FORMATO_FECHA);

  hojaFactura.getRange(`A${layout.filaResumenValores}`).setNumberFormat('0.00%');
  hojaFactura.getRange(`D${layout.filaResumenValores}`).setNumberFormat('0.00%');
  hojaFactura.getRange(`E${CONFIG.FILA_PRIMERA_LINEA}:E${layout.filaUltimaLinea}`).setNumberFormat('0.00%');
}

function restablecerLineasFactura_(hojaFactura) {
  let layout = obtenerLayoutFactura_(hojaFactura);

  const totalLineasActuales = layout.filaUltimaLinea - CONFIG.FILA_PRIMERA_LINEA + 1;
  hojaFactura.showRows(CONFIG.FILA_PRIMERA_LINEA, totalLineasActuales);
  hojaFactura.getRange(`A${CONFIG.FILA_PRIMERA_LINEA}:F${layout.filaUltimaLinea}`).clearContent();

  if (layout.filaUltimaLinea > CONFIG.FILA_ULTIMA_LINEA_BASE) {
    const filasExtra = layout.filaUltimaLinea - CONFIG.FILA_ULTIMA_LINEA_BASE;
    hojaFactura.deleteRows(CONFIG.FILA_ULTIMA_LINEA_BASE + 1, filasExtra);
  }

  hojaFactura.showRows(
    CONFIG.FILA_PRIMERA_LINEA,
    CONFIG.FILA_ULTIMA_LINEA_BASE - CONFIG.FILA_PRIMERA_LINEA + 1
  );

  hojaFactura.hideRows(
    CONFIG.FILA_PRIMERA_LINEA + 1,
    CONFIG.FILA_ULTIMA_LINEA_BASE - CONFIG.FILA_PRIMERA_LINEA
  );

  const ivaBase = [];
  for (let fila = CONFIG.FILA_PRIMERA_LINEA; fila <= CONFIG.FILA_ULTIMA_LINEA_BASE; fila++) {
    ivaBase.push([CONFIG.IVA_POR_DEFECTO]);
  }
  hojaFactura.getRange(`E${CONFIG.FILA_PRIMERA_LINEA}:E${CONFIG.FILA_ULTIMA_LINEA_BASE}`).setValues(ivaBase);
}

function obtenerLayoutFactura_(hojaFactura) {
  const filaTotalesTitulo = buscarFilaPorTextoEnColumnaA_(hojaFactura, CONFIG.TEXTO_TOTALES);

  if (!filaTotalesTitulo) {
    throw new Error(`No se encontró la fila con el texto "${CONFIG.TEXTO_TOTALES}".`);
  }

  return {
    filaTotalesTitulo: filaTotalesTitulo,
    filaUltimaLinea: filaTotalesTitulo - 1,
    filaResumenCabecera: filaTotalesTitulo + 1,
    filaResumenValores: filaTotalesTitulo + 2,
    filaDistribucionTitulo: filaTotalesTitulo + 4,
    filaDistribucionCabecera: filaTotalesTitulo + 5,
    filaDistribucionValores: filaTotalesTitulo + 6,
    filaNotas: filaTotalesTitulo + 8
  };
}

function buscarFilaPorTextoEnColumnaA_(hoja, texto) {
  const ultimaFila = hoja.getLastRow();
  const valores = hoja.getRange(1, 1, ultimaFila, 1).getDisplayValues().flat();

  const indice = valores.findIndex(v => String(v).trim() === texto);
  return indice === -1 ? null : indice + 1;
}

function hayConceptosFactura_(hojaFactura) {
  const layout = obtenerLayoutFactura_(hojaFactura);

  for (let fila = CONFIG.FILA_PRIMERA_LINEA; fila <= layout.filaUltimaLinea; fila++) {
    if (hojaFactura.isRowHiddenByUser(fila)) continue;

    const valores = hojaFactura.getRange(`A${fila}:C${fila}`).getDisplayValues()[0];
    const tieneAlgo = valores.some(v => String(v).trim() !== '');

    if (tieneAlgo) return true;
  }

  return false;
}

function normalizarNumero_(valor) {
  if (typeof valor === 'number' && !isNaN(valor)) return valor;
  return null;
}

function normalizarPorcentaje_(valor) {
  if (typeof valor !== 'number' || isNaN(valor)) return null;
  if (valor > 1) return valor / 100;
  if (valor < 0) return 0;
  return valor;
}

/* =========================
   HOJAS / TIPO DOCUMENTO
========================= */

function obtenerTipoDocumentoDeHoja_(hoja) {
  if (!hoja) return null;

  const nombre = hoja.getName();

  if (nombre === CONFIG.HOJA_FACTURA || esHojaFacturaArchivada_(hoja)) {
    return 'factura';
  }

  if (nombre === CONFIG.HOJA_PRESUPUESTO || esHojaPresupuestoArchivado_(hoja)) {
    return 'presupuesto';
  }

  return null;
}

function obtenerHojaDocumentoActiva_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva = ss.getActiveSheet();

  if (esHojaDocumentoEditable_(hojaActiva)) {
    return hojaActiva;
  }

  return null;
}

function obtenerHojaFacturaActiva_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva = ss.getActiveSheet();

  if (esHojaFacturaEditable_(hojaActiva)) {
    return hojaActiva;
  }

  return obtenerHojaPlantillaFactura_();
}

function obtenerHojaPresupuestoActiva_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva = ss.getActiveSheet();

  if (esHojaPresupuestoEditable_(hojaActiva)) {
    return hojaActiva;
  }

  return obtenerHojaPlantillaPresupuesto_();
}

function esHojaDocumentoEditable_(hoja) {
  return esHojaFacturaEditable_(hoja) || esHojaPresupuestoEditable_(hoja);
}

function esHojaFacturaEditable_(hoja) {
  if (!hoja) return false;

  const nombre = hoja.getName();
  return nombre === CONFIG.HOJA_FACTURA || esHojaFacturaArchivada_(hoja);
}

function esHojaPresupuestoEditable_(hoja) {
  if (!hoja) return false;

  const nombre = hoja.getName();
  return nombre === CONFIG.HOJA_PRESUPUESTO || esHojaPresupuestoArchivado_(hoja);
}

function repararHojaFactura_(hojaFactura) {
  aplicarFormatosFactura_(hojaFactura);
  rellenarDatosCliente_(hojaFactura);
  actualizarFechaSecundariaDocumento_(hojaFactura);
  recalcularFactura_(hojaFactura);
}

function obtenerHojaPlantillaFactura_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_FACTURA);
}

function obtenerHojaPlantillaPresupuesto_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_PRESUPUESTO);
}

function esHojaFacturaArchivada_(hoja) {
  return esHojaArchivadaSegunRegistro_(hoja, CONFIG.HOJA_REGISTRO);
}

function esHojaPresupuestoArchivado_(hoja) {
  return esHojaArchivadaSegunRegistro_(hoja, CONFIG.HOJA_REGISTRO_PRESUPUESTOS);
}

function esHojaArchivadaSegunRegistro_(hoja, nombreHojaRegistro) {
  if (!hoja) return false;

  const nombre = hoja.getName();
  const nombresReservados = [
    CONFIG.HOJA_FACTURA,
    CONFIG.HOJA_PRESUPUESTO,
    CONFIG.HOJA_BASE,
    CONFIG.HOJA_REGISTRO,
    CONFIG.HOJA_REGISTRO_PRESUPUESTOS
  ];

  if (nombresReservados.includes(nombre)) return false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistro = ss.getSheetByName(nombreHojaRegistro);
  if (!hojaRegistro) return false;

  const ultimaFila = hojaRegistro.getLastRow();
  if (ultimaFila < 2) return false;

  const nombresRegistrados = hojaRegistro
    .getRange(2, 2, ultimaFila - 1, 1)
    .getDisplayValues()
    .flat()
    .map(v => normalizarNombreHojaDocumento_(v));

  return nombresRegistrados.includes(normalizarNombreHojaDocumento_(nombre));
}

/* =========================
   REGISTROS / ARCHIVO
========================= */

function buscarFilaDocumentoEnRegistro_(hojaRegistro, numeroDocumento) {
  const ultimaFila = hojaRegistro.getLastRow();
  if (ultimaFila < 2) return null;

  const numeros = hojaRegistro
    .getRange(2, 2, ultimaFila - 1, 1)
    .getDisplayValues()
    .flat()
    .map(v => String(v).trim());

  const indice = numeros.findIndex(v => v === numeroDocumento);
  return indice === -1 ? null : indice + 2;
}

function buscarFilaFacturaEnRegistro_(hojaRegistro, numeroFactura) {
  return buscarFilaDocumentoEnRegistro_(hojaRegistro, numeroFactura);
}

function crearOActualizarHojaArchivo_(hojaOrigen, numeroDocumento) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHoja = normalizarNombreHojaDocumento_(numeroDocumento);

  const hojaExistente = ss.getSheetByName(nombreHoja);
  if (hojaExistente && hojaExistente.getSheetId() !== hojaOrigen.getSheetId()) {
    ss.deleteSheet(hojaExistente);
  }

  const hojaNueva = hojaOrigen.copyTo(ss).setName(nombreHoja);

  hojaNueva.getRange(CONFIG.RANGO_NUMERO_FACTURA).setValue(numeroDocumento);

  hojaNueva.getRange(CONFIG.RANGO_CLIENTE).setValue(
    hojaOrigen.getRange(CONFIG.RANGO_CLIENTE).getDisplayValue()
  );
  hojaNueva.getRange(CONFIG.RANGO_CIF).setValue(
    hojaOrigen.getRange(CONFIG.RANGO_CIF).getDisplayValue()
  );
  hojaNueva.getRange(CONFIG.RANGO_DIRECCION).setValue(
    hojaOrigen.getRange(CONFIG.RANGO_DIRECCION).getDisplayValue()
  );

  ss.setActiveSheet(hojaNueva);
  ss.moveActiveSheet(ss.getNumSheets());
  ss.setActiveSheet(hojaOrigen);

  return hojaNueva;
}

function normalizarNombreHojaDocumento_(texto) {
  return String(texto)
    .replace(/[\\\/\?\*\[\]:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 99);
}

function normalizarNombreHojaFactura_(texto) {
  return normalizarNombreHojaDocumento_(texto);
}

/* =========================
   PDFs / CARPETAS
========================= */

function sanitizarNombreArchivo_(texto) {
  return String(texto)
    .replace(/[\\/:*?"<>|#\[\]]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function sanitizarNombreCarpeta_(texto) {
  return String(texto || 'Sin nombre')
    .replace(/[\\/:*?"<>|#\[\]]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function obtenerOCrearCarpetaRaizFacturasCliente_() {
  return obtenerOCrearCarpetaRaizPorNombre_(CONFIG.NOMBRE_CARPETA_FACTURAS_CLIENTE);
}

function obtenerOCrearCarpetaRaizPorNombre_(nombre) {
  const nombreLimpio = sanitizarNombreCarpeta_(nombre);
  const carpetas = DriveApp.getFoldersByName(nombreLimpio);

  if (carpetas.hasNext()) {
    return carpetas.next();
  }

  return DriveApp.createFolder(nombreLimpio);
}

function obtenerOCrearSubcarpeta_(carpetaPadre, nombreSubcarpeta) {
  const nombreLimpio = sanitizarNombreCarpeta_(nombreSubcarpeta);
  const carpetas = carpetaPadre.getFoldersByName(nombreLimpio);

  if (carpetas.hasNext()) {
    return carpetas.next();
  }

  return carpetaPadre.createFolder(nombreLimpio);
}

function guardarOReemplazarArchivoEnCarpeta_(carpeta, blob, nombreArchivo) {
  const archivosExistentes = carpeta.getFilesByName(nombreArchivo);

  while (archivosExistentes.hasNext()) {
    archivosExistentes.next().setTrashed(true);
  }

  carpeta.createFile(blob.copyBlob().setName(nombreArchivo));
}

function borrarArchivoPorNombreEnArbol_(carpetaPadre, nombreArchivo) {
  const archivos = carpetaPadre.getFilesByName(nombreArchivo);
  while (archivos.hasNext()) {
    archivos.next().setTrashed(true);
  }

  const subcarpetas = carpetaPadre.getFolders();
  while (subcarpetas.hasNext()) {
    borrarArchivoPorNombreEnArbol_(subcarpetas.next(), nombreArchivo);
  }
}

function obtenerNombreMes_(fecha) {
  const meses = [
    'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
    'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
  ];
  return meses[fecha.getMonth()];
}

function obtenerNombreVisibleCarpetaMensual_() {
  return String(CONFIG.PDF_FOLDER_MENSUAL_ID || 'Facturas por mes').trim();
}

function obtenerCarpetaRaizMensualFacturas_() {
  const valorConfig = String(CONFIG.PDF_FOLDER_MENSUAL_ID || '').trim();

  if (!valorConfig) {
    throw new Error('No está configurada la carpeta raíz mensual de facturas.');
  }

  try {
    return DriveApp.getFolderById(valorConfig);
  } catch (error) {
    return obtenerOCrearCarpetaRaizPorNombre_(valorConfig);
  }
}

/* =========================
   NUMERACIÓN PRESUPUESTOS
========================= */

function generarSiguienteNumeroPresupuesto_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistro = ss.getSheetByName(CONFIG.HOJA_REGISTRO_PRESUPUESTOS);

  if (!hojaRegistro) {
    throw new Error(`No existe la hoja "${CONFIG.HOJA_REGISTRO_PRESUPUESTOS}".`);
  }

  const ultimaFila = hojaRegistro.getLastRow();
  let maximo = 0;

  if (ultimaFila >= 2) {
    const numeros = hojaRegistro
      .getRange(2, 2, ultimaFila - 1, 1)
      .getDisplayValues()
      .flat()
      .map(v => String(v).trim());

    numeros.forEach(numero => {
      const match = numero.match(/(\d+)(?!.*\d)/);
      if (match) {
        maximo = Math.max(maximo, Number(match[1]));
      }
    });
  }

  const anio = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  const correlativo = String(maximo + 1).padStart(3, '0');

  return `${CONFIG.PREFIJO_PRESUPUESTO}-${anio}-${correlativo}`;
}

function asignarNumeroPresupuestoSiFalta_(hojaPresupuesto) {
  const rango = hojaPresupuesto.getRange(CONFIG.RANGO_NUMERO_FACTURA);
  const numeroActual = rango.getDisplayValue().trim();

  if (numeroActual) {
    return numeroActual;
  }

  const nuevoNumero = generarSiguienteNumeroPresupuesto_();
  rango.setValue(nuevoNumero);
  return nuevoNumero;
}

function forzarNuevoNumeroPresupuesto_(hojaPresupuesto) {
  const nuevoNumero = generarSiguienteNumeroPresupuesto_();
  hojaPresupuesto.getRange(CONFIG.RANGO_NUMERO_FACTURA).setValue(nuevoNumero);
  return nuevoNumero;
}

/* =========================
   DESPLEGABLE CLIENTES
========================= */

function repararDropdownClientes_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaBase = ss.getSheetByName(CONFIG.HOJA_BASE);

  if (!hojaBase) return;

  const hojasObjetivo = ss.getSheets().filter(hoja => esHojaDocumentoEditable_(hoja));
  if (!hojasObjetivo.length) return;

  const ultimaFila = hojaBase.getLastRow();

  if (ultimaFila < 2) {
    hojasObjetivo.forEach(hoja => {
      hoja.getRange(CONFIG.RANGO_CLIENTE).clearDataValidations();
    });
    return;
  }

  const rangoClientes = hojaBase.getRange(2, 1, ultimaFila - 1, 1);

  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangoClientes, true)
    .setAllowInvalid(false)
    .setHelpText('Selecciona un cliente de la base de datos')
    .build();

  hojasObjetivo.forEach(hoja => {
    hoja.getRange(CONFIG.RANGO_CLIENTE).setDataValidation(regla);
  });
}

function aplicarEstiloModernoFacturas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = [
    ss.getSheetByName(CONFIG.HOJA_FACTURA),
    ss.getSheetByName(CONFIG.HOJA_PRESUPUESTO)
  ].filter(Boolean);

  hojas.forEach(hoja => aplicarEstiloVisualDocumento_(hoja));
}

function aplicarEstiloVisualDocumento_(hoja) {
  const layout = obtenerLayoutFactura_(hoja);

  const FUENTE = 'Aptos';
  const COLOR_TEXTO = '#1F2937';
  const COLOR_PRINCIPAL = '#274C77';
  const COLOR_SUAVE = '#EAF2F8';
  const COLOR_SUAVE_2 = '#F6F9FC';
  const COLOR_BORDE = '#D0D7DE';
  const COLOR_TOTAL = '#DCE6F1';
  const COLOR_TOTAL_FINAL = '#274C77';

  const ultimaFilaVisual = layout.filaNotas + 3;

  hoja.getRange(1, 1, ultimaFilaVisual, 6)
    .setFontFamily(FUENTE)
    .setFontColor(COLOR_TEXTO)
    .setVerticalAlignment('middle');

  hoja.getRange('A1:F6')
    .setFontFamily(FUENTE)
    .setFontColor(COLOR_PRINCIPAL);

  hoja.getRange('A1:F2')
    .setFontSize(14)
    .setFontWeight('bold');

  hoja.getRange('A3:F22')
    .setFontSize(10);

  hoja.getRange('A1:F22')
    .setWrap(true);

  hoja.getRange('A24:F24')
    .setBackground(COLOR_PRINCIPAL)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  hoja.getRange(`A25:F${layout.filaUltimaLinea}`)
    .setBackground(null)
    .setFontSize(10)
    .setBorder(true, true, true, true, true, true, COLOR_BORDE, SpreadsheetApp.BorderStyle.SOLID);

  hoja.getRange(`A25:A${layout.filaUltimaLinea}`).setHorizontalAlignment('left');
  hoja.getRange(`B25:F${layout.filaUltimaLinea}`).setHorizontalAlignment('center');

  hoja.getRange(`A${layout.filaTotalesTitulo}:F${layout.filaTotalesTitulo}`)
    .setBackground(COLOR_SUAVE_2)
    .setFontColor(COLOR_PRINCIPAL)
    .setFontWeight('bold');

  hoja.getRange(`A${layout.filaResumenCabecera}:F${layout.filaResumenCabecera}`)
    .setBackground(COLOR_SUAVE)
    .setFontColor(COLOR_PRINCIPAL)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, COLOR_BORDE, SpreadsheetApp.BorderStyle.SOLID);

  hoja.getRange(`A${layout.filaResumenValores}:F${layout.filaResumenValores}`)
    .setBackground('#FFFFFF')
    .setBorder(true, true, true, true, true, true, COLOR_BORDE, SpreadsheetApp.BorderStyle.SOLID)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  hoja.getRange(`F${layout.filaResumenValores}`)
    .setBackground(COLOR_TOTAL_FINAL)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11);

  hoja.getRange(`A${layout.filaResumenValores}:E${layout.filaResumenValores}`)
    .setBackground(COLOR_TOTAL);

  hoja.getRange(`A${layout.filaDistribucionTitulo}:F${layout.filaDistribucionTitulo}`)
    .setBackground(COLOR_SUAVE_2)
    .setFontColor(COLOR_PRINCIPAL)
    .setFontWeight('bold');

  hoja.getRange(`A${layout.filaDistribucionCabecera}:F${layout.filaDistribucionCabecera}`)
    .setBackground(COLOR_SUAVE)
    .setFontColor(COLOR_PRINCIPAL)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, COLOR_BORDE, SpreadsheetApp.BorderStyle.SOLID);

  hoja.getRange(`A${layout.filaDistribucionValores}:F${layout.filaDistribucionValores}`)
    .setBorder(true, true, true, true, true, true, COLOR_BORDE, SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment('center');

  hoja.getRange(`A${layout.filaNotas}:F${layout.filaNotas + 2}`)
    .setBackground(COLOR_SUAVE_2)
    .setBorder(true, true, true, true, false, false, COLOR_BORDE, SpreadsheetApp.BorderStyle.SOLID);

  hoja.getRange(CONFIG.RANGO_CLIENTE).setFontWeight('bold');
  hoja.getRange(CONFIG.RANGO_CIF).setFontWeight('bold');
  hoja.getRange(CONFIG.RANGO_DIRECCION).setFontWeight('bold');

  hoja.getRange(CONFIG.RANGO_NUMERO_FACTURA).setFontWeight('bold');
  hoja.getRange(CONFIG.RANGO_FECHA_FACTURA).setFontWeight('bold');
  hoja.getRange(CONFIG.RANGO_FECHA_VENCIMIENTO).setFontWeight('bold');

  aplicarFormatosFactura_(hoja);
}