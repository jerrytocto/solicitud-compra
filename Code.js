// Manejo de solicitudes de GET
function doGet(e) {
  console.log("ID : " + e.parameter.solicitudId);
  console.log("ESTADO : " + e.parameter.estado);
  console.log("APROBADOR : " + e.parameter.aprobadoresEmail);
  if (
    (e.parameter.solicitudId && e.parameter.estado, e.parameter.aprobadoresEmail)
  ) {

    return handleEstadoRequest(e);
  }
  var template = HtmlService.createTemplateFromFile("index");
  template.pubUrl =
    "https://script.google.com/a/macros/ahkgroup.com/s/AKfycby-Dodddv4sFV4eOrBSELXKNXoah9lkY8q_LxRO2kCu/dev";
  return template.evaluate();
}

// Manejo de solicitudes de POST
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRegistro = ss.getSheetByName("index");

  var solicitudId = generarSolicitudId(sheetRegistro);
  var totalCompra = registrarProductos(e, solicitudId, sheetRegistro);

  // Esperar 4 segundos para asegurarse de que los cambios se guarden
  Utilities.sleep(4000);

  // Enviar correo electrónico de notificación
  enviarEmail(totalCompra, solicitudId);

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
  
  console.log("Total de compra: "+ totalCompra)
  productos.forEach((producto) => {
    var subtotal = calcularSubtotal(producto.cantidad, producto.precio);

    var fila = [
      solicitudId,
      e.parameter.solicitante,
      e.parameter.email,
      e.parameter.razonCompra,
      fechaRegistro,
      e.parameter.prioridad,
      e.parameter.centroCosto,
      producto.nombre,
      producto.marca,
      producto.especificaciones,
      producto.cantidad,
      producto.precio,
      subtotal,
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

// Extrae los datos de los productos del evento POST
function obtenerDatosProductos(e) {
  var nombres = e.parameter["productNames[]"].split(",");
  var marcas = e.parameter["productBrands[]"].split(",");
  var cantidades = e.parameter["productQuantities[]"].split(",");
  var precios = e.parameter["productPrices[]"].split(",");
  var especificaciones = e.parameter["productSpecs[]"].split(",");

  return nombres.map((nombre, i) => ({
    nombre: nombre,
    marca: marcas[i] || "",
    cantidad: parseFloat(cantidades[i]) || 0,
    precio: parseFloat(precios[i]) || 0,
    especificaciones: especificaciones[i] || "",
  }));
}

// Calcula el subtotal de un producto
function calcularSubtotal(cantidad, precio) {
  if (!isNaN(cantidad) && !isNaN(precio)) {
    return cantidad * precio;
  }
  return 0;
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

  var resultado = actualizarEstado(solicitudId, estado, aprobadoresEmail);
  return ContentService.createTextOutput(resultado);
}

// Actualiza el estado de la solicitud
function actualizarEstado(solicitudId, estado, aprobadoresEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("index");
  var data = sheet.getDataRange().getValues();
  var totalCompra = costoTotalSolicitud(solicitudId);
  var solicitanteEmail = null;
  var arrayAprobadoresEmail = [];

     //Convertimos el string en un array
  if (typeof aprobadoresEmail === 'string') {
    arrayAprobadoresEmail = aprobadoresEmail.split(',');
  }

  var columnaEstado = determinarColumnaEstado(
    totalCompra,
    arrayAprobadoresEmail
  );
  
  console.log("ID de solicitud: "+ solicitudId);
  console.log("Columnas de estado(actualizarEstado): "+ columnaEstado);
  var registrosActualizados = [];

  data.forEach((row, i) => {
    if (row[0] == solicitudId) {
      console.log("Fila: "+ row[0]);
      sheet.getRange(i + 1, columnaEstado).setValue(estado);
      sheet.getRange(i + 1, columnaEstado + 1).setValue(arrayAprobadoresEmail.length===1 ? arrayAprobadoresEmail[0]:arrayAprobadoresEmail[1]);
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
      enviarCorreoAprobado(registrosActualizados, totalCompra, aprobadoresEmail, columnaEstado);
    }
    return `El estado de la solicitud ha sido actualizado a: ${estado}`;
  } else {
    return "Solicitud no encontrada.";
  }
}

// Determina la columna de estado a actualizar según el total de la compra
function determinarColumnaEstado(totalCompra, arrayAprobadoresEmail) {

  console.log("Aprobadores (determinarColumnaEstado): "+arrayAprobadoresEmail);
  console.log("Array aprobadores: "+ arrayAprobadoresEmail.length);

  if (totalCompra <= 500) {
    return 14; // Columna para jefe del área
  } else {
    //Este correo debe ser el mismo que se encuentra en la función enviar email para cuando el monto es > 500, Lïnea 324
    if (arrayAprobadoresEmail.length==1) { 
      console.log("Aprobado por el gerente de área");
      return 14; // Columna del estado para el gerente de área
    } else if (arrayAprobadoresEmail.length==2) {
      console.log("Aprobado por el gerente general");
      return 17; // Columna del estado para el gerente general  
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
      totalCompra += parseFloat(row[12]);
    }
  });

  return totalCompra;
}

// Enviar correo electrónico de aprobación
function enviarCorreoAprobado(registros, totalCompra, aprobadoresEmail, columnaEstado) {
  var destinatario;

  // Decodificar el parámetro de correos electrónicos
  aprobadoresEmail = decodeURIComponent(aprobadoresEmail);
  
  var arrayAprobadoresEmail = [];

  //Convertimos el string en un array
  if (typeof aprobadoresEmail === 'string') {
    arrayAprobadoresEmail = aprobadoresEmail.split(',');
  }

  if (totalCompra <= 500) {
    destinatario = "jesus.arias@ahkgroup.com"; //Correo del área de compras
    enviarCorreoCompras(registros, aprobadoresEmail, destinatario, totalCompra);
  } else {
    //Verificar la columna de estaddo que se ha modificado
    if (arrayAprobadoresEmail.length==1) {
      destinatario = "jerry.chuquiguanca@ahkgroup.com"; //Correo del gerente general
      aprobadoresEmail += `,${destinatario}`;
      enviarCorreoGerenteGeneral(
        registros,
        aprobadoresEmail,
        destinatario,
        totalCompra
      );

      //Correo ha sido revisado por el gerente general
    } else if (arrayAprobadoresEmail.length==2) {
      //Luego de la aprobación del gerente general se envía el correo al área de compras
      destinatario = "jesus.arias@ahkgroup.com"; //Correo del área de compras
      enviarCorreoCompras(registros, aprobadoresEmail, destinatario, totalCompra);
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
  htmlTemplate.centroDeCosto = registrosAprobados[0][6];
  htmlTemplate.tablaSolicitud = registrosAprobados;

  htmlTemplate.aprobadoresEmail = aprobadoresEmail;

  htmlTemplate.mostrarCampoAprobador = 1;
  htmlTemplate.paraAprobar = true;

  //Verificación si la compra es mayor a 1000 generar el capex y añadirlo al correo
  if (totalCompra > 1000) {
    var capex = generarCapex(totalCompra);
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
  htmlTemplate.centroDeCosto = registrosAprobados[0][6];
  htmlTemplate.tablaSolicitud = registrosAprobados;
  htmlTemplate.aprobadoresEmail = aprobadoresEmail;
  htmlTemplate.mostrarCampoAprobador = 1;
  htmlTemplate.paraAprobar = false;

  //AQUÍ ESTÁ PENDIENTE QUE SE DEBE PASAR TAMBIÉN EL CORREO DEL GERENTE DE ÁREA

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
  console.log("Datos todos: " + filteredData)

  var htmlTemplate = HtmlService.createTemplateFromFile("tablaRequisitosEmail");
  htmlTemplate.tablaSolicitud = filteredData;
  htmlTemplate.totalCompra = totalCompra ? totalCompra.toFixed(2) : "0.00";
  htmlTemplate.solicitudId = solicitudId;
  htmlTemplate.emisor = filteredData[0][1];
  htmlTemplate.razonDeCompra = filteredData[0][3];
  htmlTemplate.fechaSolicitud = filteredData[0][4];
  htmlTemplate.centroDeCosto = filteredData[0][6];
  htmlTemplate.mostrarCampoAprobador = 0;
  htmlTemplate.paraAprobar = true;

  var destinatario =
    totalCompra <= 500
      ? "jerry.chuquiguanca@ahkgroup.com" //Correo del jefe del área
      : "jerry.chuquiguanca@ahkgroup.com"; //Correo del gerente de área
  
  var aprobadoresEmail = destinatario;
  
  htmlTemplate.aprobadoresEmail = aprobadoresEmail;
  console.log("Enviar Email. Aprobadores : "+aprobadoresEmail);
  
  var html = htmlTemplate.evaluate().getContent();
  GmailApp.sendEmail(destinatario, "SOLICITUD DE COMPRA", "MENSAJE DEL EMAIL", {
    htmlBody: html,
  });
}

// Incluir archivo HTML
function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
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














