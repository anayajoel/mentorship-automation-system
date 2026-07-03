// automatizacion_mentorias.js5
// Automatización para mentorías TRHI
// Joel Anaya - 6/18/2026
// Descripción:
// Este script automatiza la generación de documentos y PDFs para las mentorías
// registradas en la hoja "Captura". Al enviar el formulario, se asigna un ID
// y una referencia únicos, se valida la información, se crea una carpeta en Drive
// para la mentoría, y se generan los documentos FOR002, FOR003 y la cuadrícula
// de evidencia de horas. También incluye funciones para aprobar/rechazar mentorías,
// regenerar documentos, y generar un memo al 35% de avance. Además, se pueden
// obtener métricas generales y por periodo para análisis de las mentorías registradas.
// NOTA: Asegúrate de configurar los IDs de las plantillas y la carpeta raíz en Drive
// antes de usar el script.
// ============================================================

// ============================================================
// GLOBAL
// ============================================================
 
const CARPETA_RAIZ_MENTORIAS = "example";
 
// ============================================================
// TRIGGER PRINCIPAL
// ============================================================
 
function onFormSubmit(e) {
 
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("captura");
 
  const lastRow = sheet.getLastRow();
 
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];
 
  const getCol = (name) => headers.indexOf(name) + 1;
 
  const COL_ID         = getCol("ID_MENTORIA");
  const COL_REFERENCIA = getCol("REFERENCIA");
  const COL_STATUS     = getCol("STATUS");
  const COL_CAP        = getCol("PDF_Capacitacion");
  const COL_PAGO       = getCol("PDF_Pagos");
  const COL_CUADRICULA = getCol("PDF_Cuadricula");
 
  // ==========================
  // ID MENTORIA
  // ==========================
 
  let nextNum = 1;

  if (lastRow > 2) {
    const ids = sheet
      .getRange(2, COL_ID, lastRow - 1)
      .getValues()
      .flat()
    .filter(id => id && id.toString().startsWith("MEN-"));

  if (ids.length > 0) {
    const nums = ids.map(id =>
      parseInt(id.toString().replace("MEN-", ""))
    );

    nextNum = Math.max(...nums) + 1;
  }
}

idMentoria = "MEN-" + nextNum.toString().padStart(4, "0");
sheet.getRange(lastRow, COL_ID).setValue(idMentoria);
 
  // ==========================
  // REFERENCIA
  // ==========================
 
  let nextRef = 1;

  if (lastRow > 2) {
    const refs = sheet
      .getRange(2, COL_REFERENCIA, lastRow - 1)
      .getValues()
      .flat()
      .filter(ref => ref && ref.toString().startsWith("TRHI-"));

  if (refs.length > 0) {
    const nums = refs.map(ref => {
      const partes = ref.toString().split("-");
      return parseInt(partes[2]);
    }).filter(n => !isNaN(n));

    if (nums.length > 0) {
      nextRef = Math.max(...nums) + 1;
    }
  }
}

  const anio = new Date().getFullYear();

  referencia =
    "TRHI-" +
    anio +
    "-" +
   nextRef.toString().padStart(3, "0");

  sheet.getRange(lastRow, COL_REFERENCIA).setValue(referencia);
 
  // ==========================
  // DATOS
  // FIX: rowData se lee DESPUÉS de escribir ID y REFERENCIA
  // para que data{} refleje la fila completa y actualizada
  // ==========================
 
  const rowData = sheet
    .getRange(lastRow, 1, 1, sheet.getLastColumn())
    .getValues()[0];
 
  const data = {};
  headers.forEach((h, i) => {
    data[h] = rowData[i];
  });
 
  const datos = {
 
    idMentoria:   idMentoria,
    referencia:   referencia,
 
    numEmpleado:
      data["Número de Empleado del Mentor"],
 
    mentor:
      data["Nombre del Mentor"],
 
    puesto:
      data["Puesto"],
 
    departamento:
      data["Departamento"],
 
    gerencia:
      data["Gerencia"],
 
    curso:
      data["Nombre del Curso"],
 
    justificacion:
      data["Justificacion"],
 
    fechaInicio:
      data["Fecha Inicio"],
 
    fechaFin:
      data["Fecha Fin"],
 
    horasTotales:
      data["Horas Totales"],
 
    participantes:
      data["Numero de Empleado - Nombre \nEjemplo ( 12345678 - Joel Anaya )"]
  };
 
  // ==========================
  // VALIDAR FECHAS
  // FIX: proteger contra fechas vacías o inválidas antes de continuar
  // ==========================
 
  if (!datos.fechaInicio || !datos.fechaFin) {
    sheet.getRange(lastRow, COL_STATUS).setValue("ERROR: FECHA VACÍA");
    return;
  }
 
  const fechaInicioObj = new Date(datos.fechaInicio);
  const fechaFinObj    = new Date(datos.fechaFin);
 
  if (isNaN(fechaInicioObj.getTime()) || isNaN(fechaFinObj.getTime())) {
    sheet.getRange(lastRow, COL_STATUS).setValue("ERROR: FECHA INVÁLIDA");
    return;
  }
 
  if (fechaFinObj < fechaInicioObj) {
    sheet.getRange(lastRow, COL_STATUS).setValue("ERROR: FECHA FIN ANTES QUE INICIO");
    return;
  }
 
  // ==========================
  // VALIDAR PARTICIPANTE
  // ==========================
 
  Logger.log("PARTICIPANTES:");
  Logger.log(datos.participantes);

  if (!validarParticipante(datos.participantes)) {

  sheet
    .getRange(lastRow, COL_STATUS)
    .setValue("ERROR PARTICIPANTE");

  return;
  }
 
  // ==========================
  // VALIDAR DUPLICADOS
  // ==========================
 
  if (
  existeDuplicado(
    sheet,
    lastRow,
    datos
  )
) {

  sheet
    .getRange(lastRow, COL_STATUS)
    .setValue("DUPLICADO");

  return;
}
  // ==========================
  // CARPETA DRIVE
  // ==========================
 
  const carpetaMentoria = obtenerOCrearCarpetaMentoria(
    referencia,
    datos.mentor,
    datos.curso
  );
 
  // ==========================
  // IDS TEMPLATES
  // ==========================
 
  const TEMPLATE_CAP        = "example";
  const TEMPLATE_PAGO       = "example";
  const TEMPLATE_CUADRICULA = "example";
 
  try {
 
    Logger.log("GENERANDO FOR002");
    const urlCap = generarDocCompleto(datos, carpetaMentoria, TEMPLATE_CAP, "FOR002_SolicitudCapacitacion");
    sheet.getRange(lastRow, COL_CAP).setValue(urlCap);
    Logger.log("FOR002 OK");
 
    Logger.log("GENERANDO FOR003");
    const urlPago = generarDocCompleto(datos, carpetaMentoria, TEMPLATE_PAGO, "FOR003_AutorizacionPago");
    sheet.getRange(lastRow, COL_PAGO).setValue(urlPago);
    Logger.log("FOR003 OK");
 
    Logger.log("GENERANDO CUADRICULA");
    const urlCuadricula = generarCuadriculaV2(datos, carpetaMentoria, TEMPLATE_CUADRICULA);
    sheet.getRange(lastRow, COL_CUADRICULA).setValue(urlCuadricula);
    Logger.log("CUADRICULA OK");
 
    sheet.getRange(lastRow, COL_STATUS).setValue("PENDIENTE REVISION");
    SpreadsheetApp.flush();
 
  } catch (error) {
 
    Logger.log(error);
    sheet.getRange(lastRow, COL_STATUS).setValue("ERROR: " + error.message);
  }
}
 
// ============================================================
// GENERAR DOCS (FOR002 / FOR003)
// ============================================================
 
function generarDocCompleto(datos, folder, templateId, prefijo) {
 
  const copy = DriveApp
    .getFileById(templateId)
    .makeCopy(prefijo + "_" + datos.referencia, folder);
 
  const doc  = DocumentApp.openById(copy.getId());
  const body = doc.getBody();
 
  const tz = Session.getScriptTimeZone();
 
  body.replaceText("{{REFERENCIA}}",     String(datos.referencia));
  body.replaceText("{{ID_MENTORIA}}",    String(datos.idMentoria));
  body.replaceText("{{NUM_EMPLEADO}}",   String(datos.numEmpleado));
  body.replaceText("{{NUM_MENTOR}}",     String(datos.numEmpleado));
  body.replaceText("{{NOMBRE_MENTOR}}",  String(datos.mentor));
  body.replaceText("{{PUESTO}}",         String(datos.puesto));
  body.replaceText("{{DEPARTAMENTO}}",   String(datos.departamento));
  body.replaceText("{{GERENCIA}}",       String(datos.gerencia));
  body.replaceText("{{CURSO}}",          String(datos.curso));
  body.replaceText("{{JUSTIFICACION}}",  String(datos.justificacion));
 
  body.replaceText("{{FECHA_INICIO}}",
    Utilities.formatDate(new Date(datos.fechaInicio), tz, "dd/MM/yyyy")
  );
 
  body.replaceText("{{FECHA_FIN}}",
    Utilities.formatDate(new Date(datos.fechaFin), tz, "dd/MM/yyyy")
  );
 
  body.replaceText("{{HORAS}}",          String(datos.horasTotales));
  body.replaceText("{{HORAS_TOTALES}}",  String(datos.horasTotales));
  body.replaceText("{{PARTICIPANTES}}",  String(datos.participantes));
 
  doc.saveAndClose();
 
  const pdf = folder.createFile(
    copy.getBlob().getAs("application/pdf")
  );
 
  DriveApp.getFileById(copy.getId()).setTrashed(true);
 
  return pdf.getUrl();
}
 
// ============================================================
// GENERAR CUADRÍCULA
// ============================================================
 
function generarCuadriculaV2(datos, folder, templateId) {
 
  const copy = DriveApp
    .getFileById(templateId)
    .makeCopy("CUADRICULA_" + datos.referencia, folder);
 
  const ss       = SpreadsheetApp.openById(copy.getId());
  const sheet    = ss.getSheetByName("Cuadrícula de Mentoría V2");
  const tz       = Session.getScriptTimeZone();
 
  // ==========================
  // ENCABEZADO
  // ==========================
 
  sheet.getRange("B4").setValue(datos.referencia);
  sheet.getRange("B5").setValue(datos.mentor);
  sheet.getRange("B6").setValue(datos.participantes);
  sheet.getRange("B7").setValue(datos.curso);
  sheet.getRange("E4").setValue(datos.departamento);
  sheet.getRange("E5").setValue(datos.gerencia);
 
  sheet.getRange("E6")
    .setValue(Utilities.formatDate(new Date(datos.fechaInicio), tz, "dd/MM/yyyy"))
    .setHorizontalAlignment("left");
 
  sheet.getRange("E7")
    .setValue(Utilities.formatDate(new Date(datos.fechaFin), tz, "dd/MM/yyyy"))
    .setHorizontalAlignment("left");
 
  sheet.getRange("E8")
    .setValue(datos.horasTotales)
    .setHorizontalAlignment("left");
 
  // ==========================
  // CÁLCULO DE SESIONES
  // ==========================
 
  const inicio = new Date(datos.fechaInicio);
  const fin    = new Date(datos.fechaFin);
 
  const diasHabiles = [];
  const cursor = new Date(inicio);
 
  while (cursor <= fin) {
    const dia = cursor.getDay();
    if (dia !== 0 && dia !== 6) {
      diasHabiles.push(new Date(cursor));
    }
    cursor.setDate(cursor.getDate() + 1);
  }
 
  const totalHoras  = Number(datos.horasTotales);
  const horasPorDia = diasHabiles.length > 0
    ? totalHoras / diasHabiles.length
    : totalHoras;
 
  const diasTexto = ["DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"];
 
  let i = 0;
 
  for (i = 0; i < diasHabiles.length && i < 15; i++) {
    const fila = 12 + i;
    sheet.getRange("A" + fila).setValue(
      Utilities.formatDate(diasHabiles[i], tz, "dd/MM/yyyy")
    );
    sheet.getRange("B" + fila).setValue(diasTexto[diasHabiles[i].getDay()]);
    sheet.getRange("C" + fila).setValue(Math.round(horasPorDia));
  }
 
  // Limpiar filas sobrantes
  for (let j = i; j < 15; j++) {
    const fila = 12 + j;
    sheet.getRange("A" + fila).setValue("");
    sheet.getRange("B" + fila).setValue("");
    sheet.getRange("C" + fila).setValue("");
  }
 
  sheet.getRange("B28").setValue(totalHoras);
  sheet.getRange("B29").setValue(diasHabiles.length);
 
  // FIX: flush ANTES de exportar a PDF para asegurar que los datos estén escritos
  SpreadsheetApp.flush();
 
  const pdf = folder.createFile(
    copy.getBlob().getAs("application/pdf")
  );
 
  pdf.setName("Cuadricula_EvidenciaHoras.pdf");
 
  DriveApp.getFileById(copy.getId()).setTrashed(true);
 
  return pdf.getUrl();
}
 
// ============================================================
// CARPETA EN DRIVE
// ============================================================
 
function obtenerOCrearCarpetaMentoria(referencia, mentor, curso) {
 
  const carpetaRaiz = DriveApp.getFolderById(CARPETA_RAIZ_MENTORIAS);
 
  const nombreCarpeta =
    referencia.replace(/\//g, "-")        + "_" +
    mentor.replace(/\s+/g, "")            + "_" +
    curso.replace(/[^a-zA-Z0-9]/g, "");
 
  const carpetas = carpetaRaiz.getFoldersByName(nombreCarpeta);
 
  if (carpetas.hasNext()) {
    return carpetas.next();
  }
 
  return carpetaRaiz.createFolder(nombreCarpeta);
}
 
// ============================================================
// VALIDAR PARTICIPANTE
// ============================================================
function validarParticipante(participante) {

  if (!participante) {
    return false;
  }

  const lineas = participante
    .toString()
    .split("\n")
    .map(l => l.trim())
    .filter(l => l.length > 0);

  if (lineas.length === 0) {
    return false;
  }

  // Formato:
  // 123456 - Juan Perez

  const patron =
    /^\d+\s*-\s*[A-Za-zÁÉÍÓÚáéíóúÑñ\s]+$/;

  const empleados = [];

  for (const linea of lineas) {

    if (!patron.test(linea)) {
      return false;
    }

    const numero =
      linea.split("-")[0].trim();

    // Evitar participantes repetidos
    if (empleados.includes(numero)) {
      return false;
    }

    empleados.push(numero);
  }

  return true;
}
 

// ============================================================
// DUPS
// ============================================================
function existeDuplicado(sheet, lastRow, datos) {

  if (lastRow <= 2) {
    return false;
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const getCol = (name) => headers.indexOf(name);

  const iNumEmpleado =
    getCol("Número de Empleado del Mentor");

  const iCurso =
    getCol("Nombre del Curso");

  const iInicio =
    getCol("Fecha Inicio");

  const iFin =
    getCol("Fecha Fin");

  if (
    iNumEmpleado === -1 ||
    iCurso === -1 ||
    iInicio === -1 ||
    iFin === -1
  ) {
    return false;
  }

  const filas = sheet
    .getRange(
      2,
      1,
      lastRow - 2,
      sheet.getLastColumn()
    )
    .getValues();

  const empleadoNuevo =
    String(datos.numEmpleado).trim();

  const cursoNuevo =
    String(datos.curso)
      .trim()
      .toLowerCase();

  const inicioNuevo =
    new Date(datos.fechaInicio)
      .toDateString();

  const finNuevo =
    new Date(datos.fechaFin)
      .toDateString();

  for (const fila of filas) {

    const empleadoFila =
      String(fila[iNumEmpleado]).trim();

    const cursoFila =
      String(fila[iCurso])
        .trim()
        .toLowerCase();

    const inicioFila =
      new Date(fila[iInicio])
        .toDateString();

    const finFila =
      new Date(fila[iFin])
        .toDateString();

    if (
      empleadoFila === empleadoNuevo &&
      cursoFila === cursoNuevo &&
      inicioFila === inicioNuevo &&
      finFila === finNuevo
    ) {

      Logger.log(
        "DUPLICADO DETECTADO"
      );

      return true;
    }
  }

  return false;
}


// ============================================================
// REGENERACION MENTORIAS
// ============================================================
function regenerarMentoria() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("captura");

  const fila =
    sheet.getActiveCell().getRow();

  if (fila <= 1) {
    SpreadsheetApp.getUi().alert(
      "Selecciona una fila de mentoría."
    );
    return;
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const rowData = sheet
    .getRange(
      fila,
      1,
      1,
      sheet.getLastColumn()
    )
    .getValues()[0];

  const data = {};

  headers.forEach((h, i) => {
    data[h] = rowData[i];
  });

  const getCol = (name) =>
    headers.indexOf(name) + 1;

  const COL_STATUS =
    getCol("STATUS");

  const COL_CAP =
    getCol("PDF_Capacitacion");

  const COL_PAGO =
    getCol("PDF_Pagos");

  const COL_CUADRICULA =
    getCol("PDF_Cuadricula");

  const datos = {

    idMentoria:
      data["ID_MENTORIA"],

    referencia:
      data["REFERENCIA"],

    numEmpleado:
      data["Número de Empleado del Mentor"],

    mentor:
      data["Nombre del Mentor"],

    puesto:
      data["Puesto"],

    departamento:
      data["Departamento"],

    gerencia:
      data["Gerencia"],

    curso:
      data["Nombre del Curso"],

    justificacion:
      data["Justificacion"],

    fechaInicio:
      data["Fecha Inicio"],

    fechaFin:
      data["Fecha Fin"],

    horasTotales:
      data["Horas Totales"],

    participantes:
      data["Numero de Empleado - Nombre \nEjemplo ( 123456 - Jesus Perez )"]
  };

  sheet
    .getRange(fila, COL_STATUS)
    .setValue("PENDIENTE REVISION");

  SpreadsheetApp.flush();

  const carpetaMentoria =
    obtenerOCrearCarpetaMentoria(
      datos.referencia,
      datos.mentor,
      datos.curso
    );

  // ==========================
  // BORRAR PDFs ANTERIORES
  // ==========================

  const archivos =
    carpetaMentoria.getFiles();

  while (archivos.hasNext()) {

    const archivo =
      archivos.next();

    if (
      archivo.getMimeType() ===
      MimeType.PDF
    ) {
      archivo.setTrashed(true);
    }
  }

  // ==========================
  // TEMPLATES
  // ==========================

  const TEMPLATE_CAP =
    "example";

  const TEMPLATE_PAGO =
    "example";

  const TEMPLATE_CUADRICULA =
    "example";

  try {

    const urlCap =
      generarDocCompleto(
        datos,
        carpetaMentoria,
        TEMPLATE_CAP,
        "FOR002_SolicitudCapacitacion"
      );

    sheet
      .getRange(fila, COL_CAP)
      .setValue(urlCap);

    const urlPago =
      generarDocCompleto(
        datos,
        carpetaMentoria,
        TEMPLATE_PAGO,
        "FOR003_AutorizacionPago"
      );

    sheet
      .getRange(fila, COL_PAGO)
      .setValue(urlPago);

    const urlCuadricula =
      generarCuadriculaV2(
        datos,
        carpetaMentoria,
        TEMPLATE_CUADRICULA
      );

    sheet
      .getRange(
        fila,
        COL_CUADRICULA
      )
      .setValue(urlCuadricula);

    sheet
      .getRange(
        fila,
        COL_STATUS
      )
      .setValue("REGENERADO");

    SpreadsheetApp.getUi().alert(
      "Mentoría regenerada correctamente."
    );

  } catch(error) {

    sheet
      .getRange(
        fila,
        COL_STATUS
      )
      .setValue(
        "ERROR: " +
        error.message
      );

    SpreadsheetApp.getUi().alert(
      error.message
    );
  }
}

// ============================================================
// GENERAR MEMO 35%
// ============================================================

function generarMemo35() {

  const TEMPLATE_MEMO35 =
    "example";

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("captura");

  const fila =
    sheet.getActiveCell().getRow();

  if (fila <= 1) {

    SpreadsheetApp.getUi().alert(
      "Selecciona una fila de mentoría."
    );

    return;
  }

  const headers = sheet
    .getRange(
      1,
      1,
      1,
      sheet.getLastColumn()
    )
    .getValues()[0];

  const rowData = sheet
    .getRange(
      fila,
      1,
      1,
      sheet.getLastColumn()
    )
    .getValues()[0];

  const data = {};

  headers.forEach((h, i) => {
    data[h] = rowData[i];
  });

  const getCol = (name) =>
    headers.indexOf(name) + 1;

  const COL_STATUS =
    getCol("STATUS");

  const COL_MEMO =
    getCol("PDF_MEMO35");

  const statusActual =
    sheet
      .getRange(
        fila,
        COL_STATUS
      )
      .getValue();

  if (
    statusActual !==
    "APROBADO"
  ) {

    SpreadsheetApp.getUi().alert(
      "La mentoría debe estar APROBADA."
    );

    return;
  }

  const datos = {

    referencia:
      data["REFERENCIA"],

    numEmpleado:
      data["Número de Empleado del Mentor"],

    mentor:
      data["Nombre del Mentor"],

    curso:
      data["Nombre del Curso"],

    fechaInicio:
      data["Fecha Inicio"],

    fechaFin:
      data["Fecha Fin"],

    horasTotales:
      data["Horas Totales"]
  };

  const carpetaMentoria =
    obtenerOCrearCarpetaMentoria(
      datos.referencia,
      datos.mentor,
      datos.curso
    );

  try {

    const urlMemo =
      generarMemo35PDF(
        datos,
        carpetaMentoria,
        TEMPLATE_MEMO35
      );

    sheet
      .getRange(
        fila,
        COL_MEMO
      )
      .setValue(urlMemo);

    sheet
      .getRange(
        fila,
        COL_STATUS
      )
      .setValue(
        "COMPLETADO"
      );

    SpreadsheetApp.getUi().alert(
      "Memo 35% generado correctamente."
    );

  } catch(error) {

    SpreadsheetApp.getUi().alert(
      error.message
    );

  }
}

// ============================================================
// PDF MEMO 35
// ============================================================

function generarMemo35PDF(
  datos,
  carpeta,
  templateId
) {

  const copia =
    DriveApp
      .getFileById(templateId)
      .makeCopy(
        "Memo35_" +
        datos.referencia,
        carpeta
      );

  const doc =
    DocumentApp.openById(
      copia.getId()
    );

  const body =
    doc.getBody();

  const fechaHoy =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );

  const periodo =
    Utilities.formatDate(
      new Date(
        datos.fechaInicio
      ),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    ) +
    " AL " +
    Utilities.formatDate(
      new Date(
        datos.fechaFin
      ),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );

  body.replaceText(
    "{{REFERENCIA}}",
    String(
      datos.referencia
    )
  );

  body.replaceText(
    "{{FECHA}}",
    fechaHoy
  );

  body.replaceText(
    "{{NOMBRE_MENTOR}}",
    String(
      datos.mentor
    )
  );

  body.replaceText(
    "{{NUM_MENTOR}}",
    String(datos.numEmpleado)
  );
  
  body.replaceText(
    "{{PERIODO}}",
    periodo
  );

  body.replaceText(
    "{{HORAS}}",
    String(
      datos.horasTotales
    )
  );

  doc.saveAndClose();

  const pdf =
    carpeta.createFile(
      copia
        .getBlob()
        .getAs(
          MimeType.PDF
        )
    );

  pdf.setName(
    "Memo35_" +
    datos.referencia +
    ".pdf"
  );

  copia.setTrashed(true);

  return pdf.getUrl();
}

// ============================================================
// BOTONES EN EL MENU DE BD_MENTORIAS
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Mentorias")
    .addItem("Aprobar Mentoría",   "aprobarMentoria")
    .addSeparator()
    .addItem("Rechazar Mentoría",  "rechazarMentoria")
    .addItem("Regenerar Mentoria", "regenerarMentoria")
    .addSeparator()
    .addItem("Generar Memo 35%", "generarMemo35")
    .addToUi();
}

// ============================================================
// APROBADO
// ============================================================
function aprobarMentoria() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("captura");

  const fila = sheet.getActiveCell().getRow();

  if (fila <= 1) {
    SpreadsheetApp.getUi().alert("Selecciona una mentoría.");
    return;
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const colStatus = headers.indexOf("STATUS") + 1;

  const statusActual = sheet.getRange(fila, colStatus).getValue();

  if (statusActual === "APROBADO") {
    SpreadsheetApp.getUi().alert("La mentoría ya está aprobada.");
    return;
  }

  sheet.getRange(fila, colStatus).setValue("APROBADO");
  SpreadsheetApp.getUi().alert("Mentoría aprobada.");
}

// ============================================================
// RECHAZADO
// ============================================================
function rechazarMentoria() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("captura");

  const fila = sheet.getActiveCell().getRow();

  if (fila <= 1) {
    SpreadsheetApp.getUi().alert("Selecciona una mentoría.");
    return;
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const colStatus = headers.indexOf("STATUS") + 1;

  sheet.getRange(fila, colStatus).setValue("RECHAZADO");
  SpreadsheetApp.getUi().alert("Mentoría rechazada.");
}

function finalizarMentoria() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Captura");

  const fila = sheet.getActiveCell().getRow();

  const headers = sheet
    .getRange(1,1,1,sheet.getLastColumn())
    .getValues()[0];

  const colStatus =
    headers.indexOf("STATUS") + 1;

  const statusActual =
    sheet.getRange(fila,colStatus).getValue();

  if (statusActual !== "APROBADO") {

    SpreadsheetApp.getUi().alert(
      "La mentoría debe estar APROBADA."
    );

    return;
  }

  sheet
    .getRange(fila,colStatus)
    .setValue("COMPLETADO");
}

// ============================================================
// CONFIGURACIÓN
// ============================================================

const CORREO_DESTINO = "example"; // prueba de correo // 

// ============================================================
// MÉTRICAS GENERALES (sin filtro de fecha)
// ============================================================

function obtenerMetricas() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Captura");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const COL_STATUS  = headers.indexOf("STATUS");
  const COL_HORAS   = headers.indexOf("Horas Totales");
  const COL_CURSO   = headers.indexOf("Nombre del Curso");
  const COL_DEPTO   = headers.indexOf("Departamento");
  const COL_GERENCIA = headers.indexOf("Gerencia");

  let totalMentorias = 0;
  let completadas    = 0;
  let pendientes     = 0;
  let horasTotales   = 0;

  const cursos       = {};
  const departamentos = {};
  const gerencias    = {};
  const estados      = {};

  for (let i = 1; i < data.length; i++) {

    totalMentorias++;

    const status   = String(data[i][COL_STATUS]);
    const horas    = Number(data[i][COL_HORAS]) || 0;
    const curso    = String(data[i][COL_CURSO]);
    const depto    = String(data[i][COL_DEPTO]);
    const gerencia = String(data[i][COL_GERENCIA]);

    horasTotales += horas;

    if (status === "COMPLETADO") completadas++;

    if (
      status === "PENDIENTE REVISION" ||
      status === "APROBADO" ||
      status === "REGENERADO"
    ) pendientes++;

    cursos[curso]            = (cursos[curso]            || 0) + 1;
    departamentos[depto]     = (departamentos[depto]     || 0) + 1;
    gerencias[gerencia]      = (gerencias[gerencia]      || 0) + 1;
    estados[status]          = (estados[status]          || 0) + 1;
  }

  return {
    totalMentorias,
    completadas,
    pendientes,
    horasTotales,
    cursosImpartidos: Object.keys(cursos).length,
    departamentos,
    gerencias,
    estados
  };
}

// ============================================================
// MÉTRICAS POR RANGO DE FECHAS
// ============================================================

function obtenerMetricasPorPeriodo(fechaInicio, fechaFin) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Captura");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const COL_FECHA    = headers.indexOf("Marca temporal");
  const COL_STATUS   = headers.indexOf("STATUS");
  const COL_HORAS    = headers.indexOf("Horas Totales");
  const COL_CURSO    = headers.indexOf("Nombre del Curso");
  const COL_DEPTO    = headers.indexOf("Departamento");
  const COL_GERENCIA = headers.indexOf("Gerencia");

  let totalMentorias = 0;
  let completadas    = 0;
  let pendientes     = 0;
  let horasTotales   = 0;

  const cursos        = {};
  const departamentos = {};
  const gerencias     = {};
  const estados       = {};

  for (let i = 1; i < data.length; i++) {

    const fechaRegistro = new Date(data[i][COL_FECHA]);

    if (fechaRegistro < fechaInicio || fechaRegistro > fechaFin) continue;

    totalMentorias++;

    const status   = String(data[i][COL_STATUS]);
    const horas    = Number(data[i][COL_HORAS]) || 0;
    const curso    = String(data[i][COL_CURSO]);
    const depto    = String(data[i][COL_DEPTO]);
    const gerencia = String(data[i][COL_GERENCIA]);

    horasTotales += horas;

    if (status === "COMPLETADO") completadas++;

    if (
      status === "PENDIENTE REVISION" ||
      status === "APROBADO" ||
      status === "REGENERADO"
    ) pendientes++;

    cursos[curso]        = (cursos[curso]        || 0) + 1;
    departamentos[depto] = (departamentos[depto] || 0) + 1;
    gerencias[gerencia]  = (gerencias[gerencia]  || 0) + 1;
    estados[status]      = (estados[status]      || 0) + 1;
  }

  return {
    totalMentorias,
    completadas,
    pendientes,
    horasTotales,
    cursosImpartidos: Object.keys(cursos).length,
    departamentos,
    gerencias,
    estados
  };
}

// ============================================================
// PERÍODOS
// ============================================================

function obtenerMetricasPrimeraQuincena() {

  const hoy = new Date();

  const inicio = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
  const fin    = new Date(hoy.getFullYear(), hoy.getMonth(), 15, 23, 59, 59);

  return obtenerMetricasPorPeriodo(inicio, fin);
}

function obtenerMetricasSegundaQuincena() {

  const hoy = new Date();

  const inicio = new Date(hoy.getFullYear(), hoy.getMonth(), 16);
  const fin    = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0, 23, 59, 59);

  return obtenerMetricasPorPeriodo(inicio, fin);
}

function obtenerMetricasMensual() {

  const hoy = new Date();

  const inicio =
    new Date(
      hoy.getFullYear(),
      hoy.getMonth(),
      1
    );

  const fin =
    new Date(
      hoy.getFullYear(),
      hoy.getMonth() + 1,
      0,
      23,
      59,
      59
    );

  return obtenerMetricasPorPeriodo(
    inicio,
    fin
  );
}

// ============================================================
// TRIGGERS
// ============================================================

function triggerQuincenal() {

  const dia = new Date().getDate();

  if (dia === 15) {
    enviarReporteQuincenal("primera");
  }

  const ultimoDia =
    new Date(
      new Date().getFullYear(),
      new Date().getMonth() + 1,
      0
    ).getDate();

  if (dia === ultimoDia) {
    enviarReporteQuincenal("segunda");
  }
}

function triggerMensual() {

  const hoy = new Date();

  const ultimoDia =
    new Date(
      hoy.getFullYear(),
      hoy.getMonth() + 1,
      0
    ).getDate();

  if (hoy.getDate() === ultimoDia) {
    enviarReporteMensual();
  }
}

// ============================================================
// ENVÍO DE REPORTES
// ============================================================

function enviarReportePrueba() {

  const m      = obtenerMetricas();
  const correo = Session.getActiveUser().getEmail();

  generarYEnviarReporte(m, "Reporte de Mentorías Internas", correo);

  SpreadsheetApp.getUi().alert("Reporte enviado a: " + correo);
}

function enviarReporteQuincenal(quincena) {

  const m = quincena === "segunda"
    ? obtenerMetricasSegundaQuincena()
    : obtenerMetricasPrimeraQuincena();

  const label = quincena === "segunda"
    ? "2da Quincena"
    : "1ra Quincena";

  generarYEnviarReporte(
    m,
    "Reporte Quincenal de Mentorías – " + label,
    CORREO_DESTINO
  );
}

function enviarReporteMensual() {

  const m = obtenerMetricasMensual();

  generarYEnviarReporte(
    m,
    "Reporte Mensual de Mentorías",
    CORREO_DESTINO
  );
}

// ============================================================
// GENERADOR DE HTML Y ENVÍO
// ============================================================

function generarYEnviarReporte(m, asunto, correo) {

  const fecha = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm"
  );

  const filaTabla = (a, b) =>
    `<tr><td>${a}</td><td>${b}</td></tr>`;

  const tablaObjeto = (obj) =>
    Object.entries(obj)
      .map(([k, v]) => filaTabla(k, v))
      .join("");

  const html = `
  <div style="font-family:Arial,sans-serif;max-width:900px;margin:auto;">

    <h1 style="
      background:#003366;
      color:white;
      padding:15px;
      text-align:center;
      border-radius:5px;
    ">${asunto}</h1>

    <p><b>Fecha de generación:</b> ${fecha}</p>

    <hr>

    <h2>Resumen General</h2>
    <table border="1" cellpadding="8" cellspacing="0"
      style="border-collapse:collapse;width:100%;">
      <tr style="background:#D9EAF7;">
        <th>Indicador</th><th>Valor</th>
      </tr>
      ${filaTabla("Mentorías Generadas",   m.totalMentorias)}
      ${filaTabla("Mentorías Completadas", m.completadas)}
      ${filaTabla("Mentorías Pendientes",  m.pendientes)}
      ${filaTabla("Horas Impartidas",      m.horasTotales)}
      ${filaTabla("Cursos Impartidos",     m.cursosImpartidos)}
    </table>

    <br>

    <h2>Mentorías por Departamento</h2>
    <table border="1" cellpadding="8" cellspacing="0"
      style="border-collapse:collapse;width:100%;">
      <tr style="background:#D9EAF7;">
        <th>Departamento</th><th>Cantidad</th>
      </tr>
      ${tablaObjeto(m.departamentos)}
    </table>

    <br>

    <h2>Mentorías por Gerencia</h2>
    <table border="1" cellpadding="8" cellspacing="0"
      style="border-collapse:collapse;width:100%;">
      <tr style="background:#D9EAF7;">
        <th>Gerencia</th><th>Cantidad</th>
      </tr>
      ${tablaObjeto(m.gerencias)}
    </table>

    <br>

    <h2>Estado de Mentorías</h2>
    <table border="1" cellpadding="8" cellspacing="0"
      style="border-collapse:collapse;width:100%;">
      <tr style="background:#D9EAF7;">
        <th>Status</th><th>Cantidad</th>
      </tr>
      ${tablaObjeto(m.estados)}
    </table>

    <br>

    <div style="color:#666;font-size:12px;text-align:center;margin-top:20px;">
      Reporte generado automáticamente por el Sistema de Mentorías.
    </div>

  </div>`;

  MailApp.sendEmail({
    to: correo,
    subject: asunto,
    htmlBody: html
  });
}