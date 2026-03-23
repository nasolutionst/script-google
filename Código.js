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

  // Intenta primero tratarlo como ID real de carpeta
  try {
    return DriveApp.getFolderById(valorConfig);
  } catch (error) {
    // Si no es un ID válido, lo tratamos como nombre de carpeta
    return obtenerOCrearCarpetaRaizPorNombre_(valorConfig);
  }
}