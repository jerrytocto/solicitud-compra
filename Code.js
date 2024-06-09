// Incluir archivo HTML
function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

// Manejo de solicitudes de GET
function doGet(e) {
  if (
    (e.parameter.solicitudId && e.parameter.estado, e.parameter.aprobadoresEmail)
  ) {

    return handleEstadoRequest(e);
  }

  var template = HtmlService.createTemplateFromFile("index");
  template.pubUrl =
    "https://script.google.com/a/macros/ahkgroup.com/s/AKfycbwm2dLfz8q4dkYCEvxU8Ic0E3SC12bC-YMXXg6_t_xtU63WeLTq8_Sv6QIhS6ynavW_CQ/exec";
  return template.evaluate();
}

// MANEJO DEL FORMULARIO
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRegistro = ss.getSheetByName("index");

  var solicitudId = generarSolicitudId(sheetRegistro);
  var totalCompra = registrarProductos(e, solicitudId, sheetRegistro);


  // Esperar 4 segundos para asegurarse de que los cambios se guarden
  Utilities.sleep(4000);

  // Enviar correo electrónico de notificación
  enviarEmail(totalCompra, solicitudId);


  //Mostrar mensaje cada vez que se envía el formulario
  return ContentService.createTextOutput(
    "TU SOLICITUD DE COMPRA ESTÁ SIENDO PROCESADA. GRACIAS"
  );
}


// Genera un ID único para la solicitud
function generarSolicitudId(sheetRegistro) {
  var lastRow = sheetRegistro.getLastRow();
  if (lastRow > 1) {
    var lastId = sheetRegistro.getRange(lastRow, 1).getValue();
    return lastId + 1;
  }
  return 1000;
}

// Registra los productos en la hoja de cálculo
function registrarProductos(e, solicitudId, sheetRegistro) {
  var fechaRegistro = new Date();
  var totalCompra = 0;
  var productos = obtenerDatosProductos(e);

  productos.forEach((producto) => {
    var subtotal = calcularSubtotal(producto.cantidad, producto.precio);
    totalCompra += subtotal;
  });

  productos.forEach((producto) => {
    var subtotal = calcularSubtotal(producto.cantidad, producto.precio);

    var fila = [
      solicitudId,
      e.parameter.solicitante,
      e.parameter.email,
      e.parameter.razonCompra,
      fechaRegistro,
      e.parameter.prioridad,
      e.parameter.justificacion,
      producto.nombre,
      producto.marca,
      producto.especificaciones,
      producto.centroCosto,
      producto.cantidad,
      producto.precio,
      subtotal,
      e.parameter.observaciones
    ];

    if (totalCompra <= 500) {
      fila.push("Pendiente", "", "");
    } else {
      fila.push("Pendiente", "", "", "Pendiente", "", "");
    }

    sheetRegistro.appendRow(fila);
  });

  return totalCompra;
}


// Maneja la solicitud de actualización de estado
function handleEstadoRequest(e) {
  var solicitudId = parseInt(e.parameter.solicitudId, 10);
  var estado = e.parameter.estado;
  var aprobadoresEmail = e.parameter.aprobadoresEmail;

  if (!solicitudId || !estado || !aprobadoresEmail) {
    return ContentService.createTextOutput(
      "Solicitud ID o estado faltante o aprobadorEmail."
    );
  }

  //Verificar el estado de la solicitud en la hoja de google sheet
  var comprobacionEstado = obtenerEstadoDeSolicitud(solicitudId,aprobadoresEmail);

  console.log("Verificacion de estado de solicitud: "+ comprobacionEstado);
  if (comprobacionEstado === "Aprobado") {
    return ContentService.createTextOutput(
      "La solicitud ya ha sido aprobada."
    );
  }

  if (comprobacionEstado === "Desaprobado") {
    return ContentService.createTextOutput(
      "La solicitud ya ha sido desaprobada."
    );
  }

  var resultado = actualizarEstado(solicitudId, estado, aprobadoresEmail);
  return ContentService.createTextOutput(resultado);
}

//Función que verifica el estado del correo antes de realizar una actualización en google sheet
function obtenerEstadoDeSolicitud(solicitudId, aprobadoresEmail) {
  var arrayAprobadoresEmail = [];
  var totalCompra = costoTotalSolicitud(solicitudId);



  // Decodificar el parámetro de correos electrónicos
  aprobadoresEmail = decodeURIComponent(aprobadoresEmail);

  //Convertimos el string en un array
  arrayAprobadoresEmail = convertOfStringToArray(aprobadoresEmail);

  console.log("Longitud de los solicitantes: "+ arrayAprobadoresEmail.length);
  console.log("typo de los solicitantes: "+ typeof arrayAprobadoresEmail.length);
  console.log("Los solicitantes: "+ arrayAprobadoresEmail);
  var columnaEstado = determinarColumnaEstado(
    totalCompra,
    arrayAprobadoresEmail
  );

  console.log("Columna de estado: "+ columnaEstado);

  var registros = obtenerUltimosRegistros(solicitudId);
  console.log("Valor del estado: "+ registros[0][columnaEstado-1]);

  return registros[0][columnaEstado-1];
}

// Extrae los datos de los productos del evento POST
function obtenerDatosProductos(e) {
  var nombres = e.parameter["productNames[]"].split(",");
  var marcas = e.parameter["productBrands[]"].split(",");
  var cantidades = e.parameter["productQuantities[]"].split(",");
  var precios = e.parameter["productPrices[]"].split(",");
  var especificaciones = e.parameter["productSpecs[]"].split(",");
  var centroCostos = e.parameter["productCentroCostos[]"].split(",");

  return nombres.map((nombre, i) => ({
    nombre: nombre,
    marca: marcas[i] || "",
    cantidad: parseFloat(cantidades[i]) || 0,
    precio: parseFloat(precios[i]) || 0,
    especificaciones: especificaciones[i] || "",
    centroCosto: centroCostos[i] || ""
  }));
}

// Calcula el subtotal de un producto
function calcularSubtotal(cantidad, precio) {
  if (!isNaN(cantidad) && !isNaN(precio)) {
    return cantidad * precio;
  }
  return 0;
}

// Actualiza el estado de la solicitud
function actualizarEstado(solicitudId, estado, aprobadoresEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("index");
  var data = sheet.getDataRange().getValues();
  var totalCompra = costoTotalSolicitud(solicitudId);
  var solicitanteEmail = null;
  var arrayAprobadoresEmail = [];

  // Decodificar el parámetro de correos electrónicos
  aprobadoresEmail = decodeURIComponent(aprobadoresEmail);

  //Convertimos el string en un array
  arrayAprobadoresEmail = convertOfStringToArray(aprobadoresEmail);

  var columnaEstado = determinarColumnaEstado(
    totalCompra,
    arrayAprobadoresEmail
  );

  var registrosActualizados = [];

  data.forEach((row, i) => {
    if (row[0] == solicitudId) {
      sheet.getRange(i + 1, columnaEstado).setValue(estado);
      sheet.getRange(i + 1, columnaEstado + 1).setValue(arrayAprobadoresEmail.length === 1 ? arrayAprobadoresEmail[0] : arrayAprobadoresEmail[1]);
      sheet.getRange(i + 1, columnaEstado + 2).setValue(new Date());
      registrosActualizados.push(row);
      solicitanteEmail = row[2];
    }
  });

  if (registrosActualizados.length > 0) {
    if (solicitanteEmail) {
      // Enviar correo para notificación al emisor del correo
      enviarCorreoRemitente(solicitanteEmail, solicitudId, estado);
    }
    if (estado === "Aprobado") {
      enviarCorreoAprobado(registrosActualizados, totalCompra, aprobadoresEmail);
    }
    return `El estado de la solicitud ha sido actualizado a: ${estado}`;
  } else {
    return "Solicitud no encontrada.";
  }
}

// Determina la columna de estado a actualizar según el total de la compra
function determinarColumnaEstado(totalCompra, arrayAprobadoresEmail) {

  if (totalCompra <= 500) {
    return 16; // Columna para jefe del área
  } else {
    if (arrayAprobadoresEmail.length == 1) {
      return 16; // Columna del estado para el gerente de área
    } else if (arrayAprobadoresEmail.length == 2) {
      return 19; // Columna del estado para el gerente general  
    }
  }
}

// Calcula el costo total de la solicitud
function costoTotalSolicitud(solicitudId) {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("index");
  var data = hoja.getDataRange().getValues();
  var totalCompra = 0;

  data.forEach((row) => {
    if (row[0] == solicitudId) {
      totalCompra += parseFloat(row[13]);
    }
  });

  return totalCompra;
}

// Enviar correo electrónico de aprobación
function enviarCorreoAprobado(registros, totalCompra, aprobadoresEmail) {
  var destinatario;

  // Decodificar el parámetro de correos electrónicos
  aprobadoresEmail = decodeURIComponent(aprobadoresEmail);

  var arrayAprobadoresEmail = [];

  //Convertimos el string en un array
  arrayAprobadoresEmail = convertOfStringToArray(aprobadoresEmail); 

  //Convertir en nombre al correo del aprobador
  var nombreAprobador = convertEmailANombre(arrayAprobadoresEmail);

  if (totalCompra <= 500) {
    destinatario = "jerrytocto@gmail.com"; //Correo del área de compras
    var nombreCargoAprobador = nombreAprobador + "- (Jefe de IT)"
    enviarCorreoCompras(registros, nombreCargoAprobador, destinatario, totalCompra);
  } else {
    //Verificar la columna de estaddo que se ha modificado
    if (arrayAprobadoresEmail.length == 1) {
      destinatario = "jerry.chuquiguanca@ahkgroup.com"; //Correo del gerente general
      arrayAprobadoresEmail.push(destinatario);
      aprobadoresEmail = arrayAprobadoresEmail.join(',');
      enviarCorreoGerenteGeneral(
        registros,
        aprobadoresEmail,
        destinatario,
        totalCompra
      );

      //Correo ha sido revisado y aprobado por el gerente general
    } else if (arrayAprobadoresEmail.length == 2) {
      //Luego de la aprobación del gerente general se envía el correo al área de compras
      destinatario = "jerrytocto@gmail.com"; //Correo del área de compras
      var arrayAprobadores = convertOfStringToArray(nombreAprobador);
      var nombreAreaGA = arrayAprobadores[0] + "- (Gerente GAF)";
      var nombreAreaGG = arrayAprobadores[1] + "- (Gerente General)";

      var nombreCargoAprobadores = [nombreAreaGA, nombreAreaGG];
      // Generar el CAPEX firmado
      var idSolicitud = registros[0][0];
      var capexFirmado = generateCapex(totalCompra, registros, nombreCargoAprobadores, true);
      enviarCorreoGerenteGeneral
      agregarLinkCapexFirmado(idSolicitud, capexFirmado.link);

      enviarCorreoCompras(
        registros,
        (nombreAreaGA + ", " + nombreAreaGG),
        destinatario,
        totalCompra
      );
    }
  }
}

// Función para añadir nuevo link del capex firmado
function agregarLinkCapexFirmado(idSolicitud, nuevoCapexLink) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("index");
  var data = sheet.getDataRange().getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === idSolicitud) {
      sheet.getRange(i + 1, 23).setValue(nuevoCapexLink);
    }
  }
}


// Enviar correo al gerente general
function enviarCorreoGerenteGeneral(
  registrosAprobados,
  aprobadoresEmail,
  destinatario,
  totalCompra
) {
  var htmlTemplate = HtmlService.createTemplateFromFile("tablaRequisitosEmail");

  htmlTemplate.solicitudId = registrosAprobados[0][0];
  htmlTemplate.emisor = registrosAprobados[0][1];
  htmlTemplate.razonDeCompra = registrosAprobados[0][3];
  htmlTemplate.fechaSolicitud = registrosAprobados[0][4];
  htmlTemplate.justificacion = registrosAprobados[0][6];
  htmlTemplate.centroDeCosto = registrosAprobados[0][10];
  htmlTemplate.observaciones = registrosAprobados[0][14];
  htmlTemplate.tablaSolicitud = registrosAprobados;

  //MÉTODO PARA TRANSFORMAR EL STRING EN UN ARRAY
  var arrayAprobadores = convertOfStringToArray(aprobadoresEmail);

  var nombresAprobadores = convertEmailANombre([arrayAprobadores[0]]);


  htmlTemplate.aprobadoresEmail = aprobadoresEmail;
  htmlTemplate.nombreCargoAprobador = nombresAprobadores + "- (Jefe de GAF)";

  htmlTemplate.mostrarCampoAprobador = 1;
  htmlTemplate.paraAprobar = true;

  //Verificación si la compra es mayor a 1000 generar el capex y añadirlo al correo
  if (totalCompra > 1000) {

    var capex = generateCapex(totalCompra, registrosAprobados, [(nombresAprobadores += "- (Jefe de GAF)")], false); // Llamar a generateCapex
    htmlTemplate.totalCompra = totalCompra.toFixed(2);

    var html = htmlTemplate.evaluate().getContent();

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("index");
    var data = sheet.getDataRange().getValues(); // Obtener todos los datos de la hoja

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(htmlTemplate.solicitudId)) { // Si el ID de la solicitud coincide
        //var lastColumn = data[i].length + 1; // Obtener la última columna
        sheet.getRange(i + 1, 22).setValue(capex.link); // Establecer el valor de la última columna al enlace
      }
    }

    GmailApp.sendEmail(
      destinatario,
      "Nueva Solicitud de Compra para Aprobar",
      "Nueva solicitud de compra aprobada.",
      {
        htmlBody: html,
        attachments: [capex.pdf]
      }
    );


  } else {
    //Solo se envía el correo con los datos de la solicitud
    htmlTemplate.totalCompra = totalCompra.toFixed(2);

    var html = htmlTemplate.evaluate().getContent();

    GmailApp.sendEmail(
      destinatario,
      "Nueva Solicitud de Compra para Aprobar",
      "Nueva solicitud de compra aprobada.",
      { htmlBody: html }
    );
  }
}

// Enviar correo de notificación al remitente
function enviarCorreoRemitente(email, solicitudId, estado) {
  var subject = `Estado de tu Solicitud de Compra #${solicitudId}`;
  var body = `Tu solicitud de compra con ID ${solicitudId} ha sido ${estado}. GRACIAS`;

  GmailApp.sendEmail(email, subject, body);
}

// Enviar correo al área de compras
function enviarCorreoCompras(
  registrosAprobados,
  aprobadoresEmail,
  destinatario,
  totalCompra
) {
  var htmlTemplate = HtmlService.createTemplateFromFile("tablaRequisitosEmail");

  htmlTemplate.solicitudId = registrosAprobados[0][0];
  htmlTemplate.emisor = registrosAprobados[0][1];
  htmlTemplate.razonDeCompra = registrosAprobados[0][3];
  htmlTemplate.fechaSolicitud = registrosAprobados[0][4];
  htmlTemplate.justificacion = registrosAprobados[0][6];
  htmlTemplate.centroDeCosto = registrosAprobados[0][10];
  htmlTemplate.observaciones = registrosAprobados[0][14];
  htmlTemplate.tablaSolicitud = registrosAprobados;
  htmlTemplate.aprobadoresEmail = aprobadoresEmail;
  htmlTemplate.mostrarCampoAprobador = 1;
  htmlTemplate.paraAprobar = false;

  htmlTemplate.nombreCargoAprobador = aprobadoresEmail;

  htmlTemplate.totalCompra = totalCompra.toFixed(2);

  var html = htmlTemplate.evaluate().getContent();

  GmailApp.sendEmail(
    destinatario,
    "Nueva Solicitud de Compra Aprobada",
    "Nueva solicitud de compra aprobada.",
    { htmlBody: html }
  );
}

// Función para enviar email para su aprobación
function enviarEmail(totalCompra, solicitudId) {
  var filteredData = obtenerUltimosRegistros(solicitudId);

  var htmlTemplate = HtmlService.createTemplateFromFile("tablaRequisitosEmail");
  htmlTemplate.tablaSolicitud = filteredData;
  htmlTemplate.totalCompra = totalCompra ? totalCompra.toFixed(2) : "0.00";
  htmlTemplate.solicitudId = solicitudId;
  htmlTemplate.emisor = filteredData[0][1];
  htmlTemplate.razonDeCompra = filteredData[0][3];
  htmlTemplate.fechaSolicitud = filteredData[0][4];
  htmlTemplate.justificacion = filteredData[0][6];
  htmlTemplate.centroDeCosto = filteredData[0][10];
  htmlTemplate.observaciones = registrosAprobados[0][14];
  htmlTemplate.mostrarCampoAprobador = 0;
  htmlTemplate.paraAprobar = true;
  htmlTemplate.nombreCargoAprobador = '';

  var destinatario =
    totalCompra <= 500
      ? "jerry.chuquiguanca@ahkgroup.com" //Correo del jefe del área
      : "jerry.chuquiguanca@ahkgroup.com"; //Correo del gerente de área

  var aprobadoresEmail = destinatario;

  htmlTemplate.aprobadoresEmail = aprobadoresEmail;

  var html = htmlTemplate.evaluate().getContent();
  GmailApp.sendEmail(destinatario, "SOLICITUD DE COMPRA", "MENSAJE DEL EMAIL", {
    htmlBody: html,
  });
}

// Obtener los últimos registros de una solicitud
function obtenerUltimosRegistros(solicitudId) {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("index");
  var data = hoja.getDataRange().getValues();
  var totalCompra = costoTotalSolicitud(solicitudId);
  var filteredData = data.filter((row) => row[0] === solicitudId);
  return filteredData;
}

// FUNCIÓN PARA TRANSFORMAR EL NOMBRE DE UN APROBADOR
function convertEmailANombre(aprobadoresEmail) {
  var nombresAprobadores = '';

  for (var i = 0; i < aprobadoresEmail.length; i++) {
    var email = aprobadoresEmail[i];
    var partes = email.split('@')[0].split('.');
    var nombre = partes[0];
    var apellido = partes[1];

    // Capitalizar la primera letra del nombre y apellido
    nombre = nombre.charAt(0).toUpperCase() + nombre.slice(1);
    apellido = apellido.charAt(0).toUpperCase() + apellido.slice(1);

    nombresAprobadores += nombre + ' ' + apellido;

    // Añadir una coma y espacio si no es el último elemento
    if (i < aprobadoresEmail.length - 1) {
      nombresAprobadores += ', ';
    }
  }

  return nombresAprobadores;
}

//FUNCIÓN PARA CONVERTIR UN STRING A UN ARRAY
function convertOfStringToArray(aprobadoresEmail) {
  var arrayAprobadoresEmail = [];

  //Convertimos el string en un array
  if (typeof aprobadoresEmail === 'string') {
    arrayAprobadoresEmail = aprobadoresEmail.split(',');
  }
  return arrayAprobadoresEmail;
}











