function doGet(e) {

  if (e.parameter.solicitudId && e.parameter.estado && e.parameter.aprobadorEmail) {
    return handleEstadoRequest(e);
  } else if (e.parameter.solicitudId && e.parameter.estado) {
    return handleEstadoRequest(e);
  }

  var template = HtmlService.createTemplateFromFile('index');
  template.pubUrl = "https://script.google.com/a/macros/ahkgroup.com/s/AKfycby-Dodddv4sFV4eOrBSELXKNXoah9lkY8q_LxRO2kCu/dev";
  var output = template.evaluate();
  return output;
}

//Carga o incluye tanto el archivo css y js en el archivo index
function include(fileName) {

  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

//Función que se llamará cada vez que se quiera registrar un producto, recibe como parámetro un objeto evento
function doPost(e) {

  // abre el libro de google sheet 
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // De todo el libro toma la hoja con nombre index
  var sheetRegistro = ss.getSheetByName('index');

  //Acceso a los datos principal del formulario
  var fechaRegistro = new Date();
  var solicitante = e.parameter.solicitante;
  var email = e.parameter.email;
  var razonCompra = e.parameter.razonCompra;
  var prioridad = e.parameter.prioridad;
  var centroCosto = e.parameter.centroCosto;

  //Acceso para los datos de la tabla de productos a solicitar
  var productNames = e.parameter['productNames[]'].split(',');
  var productBrands = e.parameter['productBrands[]'].split(',');
  var productQuantities = e.parameter['productQuantities[]'].split(',');
  var productPrices = e.parameter['productPrices[]'].split(',');
  var productSpecs = e.parameter['productSpecs[]'].split(',');

  // Obtiene el próximo ID único
  var lastRow = sheetRegistro.getLastRow();
  var solicitudId = 1000;

  if (lastRow > 1) { // La primera fila es para el encabezado
    var lastId = sheetRegistro.getRange(lastRow, 1).getValue(); // Id se encuentra en la primera columna de la hoja
    solicitudId = lastId + 1;
  }

  // Variable para almacenar el total de la compra
  var totalCompra = 0;

  // Registra los productos en la hoja de cálculo
  for (var i = 0; i < productNames.length; i++) {
    var cantidad = parseFloat(productQuantities[i]);
    var precio = parseFloat(productPrices[i]);
    var subtotal = 0;
    if (!isNaN(cantidad) && !isNaN(precio)) {
      subtotal = cantidad * precio;
      totalCompra += subtotal;
    }

    // Verificamos si hay especificaciones para este producto
    var specs = "";
    if (productSpecs.length > i) {
      specs = productSpecs[i];
    }

    if (totalCompra <= 500) { //Área de It
      sheetRegistro.appendRow([solicitudId, solicitante, email, razonCompra, fechaRegistro, prioridad, centroCosto, productNames[i], productBrands[i], specs, productQuantities[i], productPrices[i], subtotal, "Pendiente", ""]);

    } else if (totalCompra <= 100) { //Gerencia de Administración y finanzas
      sheetRegistro.appendRow([solicitudId, solicitante, email, razonCompra, fechaRegistro, prioridad, centroCosto, productNames[i], productBrands[i], specs, productQuantities[i], productPrices[i], subtotal, "", "", "Pendiente", "", "Pendiente", ""]);

    } else {  //Gerencia de administración y finanzas con capex

    }

  }

  // Espera un 3 segundos para asegurarse de que los cambios se guarden
  Utilities.sleep(4000);

  //Llamamos a la función enviar email
  enviarEmail(totalCompra, solicitudId);

  // Devuelve una respuesta al cliente si es necesario
  return ContentService.createTextOutput('TU SOLICITUD DE COMPRA ESTÁ SIENDO PROCESADA.GRACIAS');
}

//Función para obtener los últimos registros
function obtenerUltimosRegistros(solicitudId) {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName('index');
  var data = hoja.getDataRange().getValues();
  var totalCompra = costoTotalSolicitud(solicitudId);
  var filteredData = data.filter(function (row, index) {
    return index === 0 || row[0] === solicitudId;
  }).map(function (row) {
    return row.slice(0, (totalCompra > 500) ? (row.length - 2) : (row.length - 1));
  });
  return filteredData;
}

//Obtener el total de la compra
function costoTotalSolicitud(solicitudId) {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName('index');
  var data = hoja.getDataRange().getValues();
  var totalCompra = 0;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == solicitudId) {
      totalCompra += parseFloat(data[i][12]);
    }
  }
  return totalCompra;
}

//Función para enviar email para su aprobación 
function enviarEmail(totalCompra, solicitudId) {

  //Obtener los registros filtrados
  var filteredData = obtenerUltimosRegistros(solicitudId);

  // Carga la plantilla tablaRequisitosEmail y la almacena en html
  var htmlTemplate = HtmlService.createTemplateFromFile("tablaRequisitosEmail");

  // Cargamos la tabla del htmlTemplate con los datos filtrados
  htmlTemplate.tablaSolicitud = filteredData;

  //Se pasa como parámetro el total de compras
  htmlTemplate.totalCompra = totalCompra ? totalCompra.toFixed(2) : '0.00';

  htmlTemplate.solicitudId = solicitudId; // Pasar el ID de la solicitud al template
  htmlTemplate.emisor = filteredData[1][1];
  htmlTemplate.razonDeCompra = filteredData[1][3];
  htmlTemplate.fechaSolicitud = filteredData[1][4];
  htmlTemplate.centroDeCosto = filteredData[1][6];
  htmlTemplate.aprobadorEmail = null;
  htmlTemplate.esAprobacion = false;
  var html = htmlTemplate.evaluate().getContent();

  //Verifico el monto de la solicitud para luego enviar el email de la solicitud de compra
  var destinatario = totalCompra <= 500 ? "jerry.chuquiguanca@ahkgroup.com" : "jesus.arias@ahkgroup.com";
  GmailApp.sendEmail(destinatario, "SOLICITUD DE COMPRA", "MENSAJE DEL EMAIL", { htmlBody: html });
}

//Función para actualizar el estado de la solicitud en google sheet 
function actualizarEstado(solicitudId, estado, aprobadorEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('index');
  var data = sheet.getDataRange().getValues();
  var registrosActualizados = [];
  var solicitanteEmail = null;

  //Verificar el monto de la aprobación para saber qué columna de aprobación se tiene que actualizar
  var totalCompra = costoTotalSolicitud(solicitudId);


  //ACTUALIZAR ESTADO CUANDO LA COMPRA ES MENOR O IGUAL A 500
  if (totalCompra <= 500) { //Debería aprobarse por el jefe del área, aquí se debe acualizar el estado y además agregar el email del aprobador
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == solicitudId) {
        sheet.getRange(i + 1, (data[0].length - 8)).setValue(estado); // Actualiza el estado
        registrosActualizados.push(data[i]);

        // Obtén el email del solicitante (tomará la columna 3 de la hoja que será el correo del remitente)
        solicitanteEmail = data[i][2];
      }
    }

    if (registrosActualizados.length > 0) {
      // Enviar correo al remitente con el estado de la solicitud
      if (solicitanteEmail) {
        enviarCorreoRemitente(solicitanteEmail, solicitudId, estado);
      }

      // Enviar correo al área de compras si el estado es "Aprobado"
      if (estado === "Aprobado") {
        datosParaCompraEmail = obtenerUltimosRegistros(solicitudId);
        enviarCorreoCompras(datosParaCompraEmail, aprobadorEmail);
      }
      return 'El estado de la solicitud ha sido actualizado a: ' + estado;
    } else {
      return 'Solicitud no encontrada.';
    }


    //ACTUALIZAR ESTADO PARA CUANDO LA COMPRA SEA MENOR A 1000
  } else if (totalCompra <= 1000) {//Debería aprobarse por el jefe del área, aquí se debe acualizar el estado y además agregar el email del aprobador
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == solicitudId) {
        sheet.getRange(i + 1, (data[0].length - 5)).setValue(estado); // Actualiza el estado
        registrosActualizados.push(data[i]);

        // Obtén el email del solicitante (tomará la columna 3 de la hoja que será el correo del remitente)
        solicitanteEmail = data[i][2];
      }
    }

    //Después de aprobarse por el gerente respectivo, se debe enviar el formulario hacia el gerente general
    if (registrosActualizados.length > 0) {

      // Enviar correo al remitente con el estado de la solicitud
      if (solicitanteEmail) {
        enviarCorreoRemitente(solicitanteEmail, solicitudId, estado);
      }

      // SOLO SI EL ESTADO ES APROBADO DEBE ENVIARSE AL GERENTE GENERAL
      if (estado === "Aprobado") {
        datosParaCompraEmail = obtenerUltimosRegistros(solicitudId);
        enviarCorreoGG(datosParaCompraEmail, aprobadorEmail);

        
      }
      return 'El estado de la solicitud ha sido actualizado a: ' + estado;
    } else {
      return 'Solicitud no encontrada.';
    }
  }
}

//Función para recepcionar la respuesta del correo
function handleEstadoRequest(e) {
  var solicitudId = parseInt(e.parameter.solicitudId, 10);
  var estado = e.parameter.estado;
  var aprobadorEmail = e.parameter.aprobadorEmail;

  if (!solicitudId || !estado) {
    return ContentService.createTextOutput('Solicitud ID o estado faltante.');
  }

  var resultado = actualizarEstado(solicitudId, estado, aprobadorEmail);
  return ContentService.createTextOutput(resultado);
}

//Función para enviar el correo de respuesta hacia el remitente
function enviarCorreoRemitente(email, solicitudId, estado) {
  var subject = 'Estado de tu Solicitud de Compra #' + solicitudId;
  var body = 'Tu solicitud de compra con ID ' + solicitudId + ' ha sido ' + estado + '.GRACIAS';

  GmailApp.sendEmail(email, subject, body);
}

//Función para enviar la solicitud al área de compras en el caso la solicitud sea aprobada
function enviarCorreoCompras(registrosAprobados, aprobadorEmail) {
  var htmlTemplate = HtmlService.createTemplateFromFile("tablaRequisitosEmail");

  // Asignar los datos a la plantilla
  htmlTemplate.solicitudId = registrosAprobados[1][0]; // Pasar el ID de la solicitud al template
  htmlTemplate.emisor = registrosAprobados[1][1];
  htmlTemplate.razonDeCompra = registrosAprobados[1][3];
  htmlTemplate.fechaSolicitud = registrosAprobados[1][4];
  htmlTemplate.centroDeCosto = registrosAprobados[1][6];
  htmlTemplate.tablaSolicitud = registrosAprobados;
  htmlTemplate.aprobadorEmail = aprobadorEmail;
  htmlTemplate.esAprobacion = true;
// Calcula el total de la compra
  var totalCompra = 0;
  for (var i = 0; i < registrosAprobados.length; i++) {
    totalCompra += parseFloat(registrosAprobados[i][12]); // Suponiendo que el subtotal está en la columna 11
  }

  htmlTemplate.totalCompra = totalCompra;

  var html = htmlTemplate.evaluate().getContent();

  // Enviar correo al área de compras
  GmailApp.sendEmail(
    "jesus.arias@ahkgroup.com", // CAMBIAR AL CORREO DE COMPRAS
    "Nueva Solicitud de Compra Aprobada",
    "Nueva solicitud de compra aprobada.",
    { htmlBody: html }
  );
}

//Enviar correo a gerente general
function enviarCorreoGG(registrosAprobados, aprobadorEmail) {
  var htmlTemplate = HtmlService.createTemplateFromFile("tablaRequisitosEmail");

  // Asignar los datos a la plantilla
  htmlTemplate.solicitudId = registrosAprobados[1][0]; // Pasar el ID de la solicitud al template
  htmlTemplate.emisor = registrosAprobados[1][1];
  htmlTemplate.razonDeCompra = registrosAprobados[1][3];
  htmlTemplate.fechaSolicitud = registrosAprobados[1][4];
  htmlTemplate.centroDeCosto = registrosAprobados[1][6];
  htmlTemplate.tablaSolicitud = registrosAprobados;
  htmlTemplate.aprobadorEmail = aprobadorEmail;
  htmlTemplate.esAprobacion = false;

  // Calcula el total de la compra
  var totalCompra = 0;
  for (var i = 0; i < registrosAprobados.length; i++) {
    totalCompra += parseFloat(registrosAprobados[i][12]); // Suponiendo que el subtotal está en la columna 11
  }

  htmlTemplate.totalCompra = totalCompra;

  var html = htmlTemplate.evaluate().getContent();

  // Enviar correo al área de compras
  GmailApp.sendEmail(
    "jesus.arias@ahkgroup.com", // CAMBIAR AL CORREO DE COMPRAS
    "Nueva Solicitud de Compra Aprobada",
    "Nueva solicitud de compra aprobada.",
    { htmlBody: html }
  );
}













