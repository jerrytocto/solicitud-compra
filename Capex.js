function generateCapex(totalCompra, registrosAprobados, aprobadoresEmail) {


  //Identificadores de los documentos a usar (id)
  var capexPlantillaId = "1f4gccifvzDY23aWDhEaeEHSNzEEgW-vjdVIQE00HsPk"; // Id de la plantilla de capex
  var pdfId = "1T_vTy4BVj3ypQbes5jMm4yWYOaWmacF3";      // Id de la carpeta de PDFs
  var tempId = "1PN_mQmT_RoOZJCRwhPd0xr8zkeNrNeIP";     // Id de la carpeta temporal 

  //Para poder realizar los cambios en Google Docs 
  var capexPlantilla = DriveApp.getFileById(capexPlantillaId);
  var carpetaPdf = DriveApp.getFolderById(pdfId);
  var carpetaTemp = DriveApp.getFolderById(tempId);

  // Hacer una copia de la plantilla en la carpeta temporal
  var copiaPlantilla = capexPlantilla.makeCopy(carpetaTemp);
  var copiaId = copiaPlantilla.getId();
  var doc = DocumentApp.openById(copiaId);

  // Datos generales (asumiendo que los datos generales están en la primera fila)
  var solicitante = registrosAprobados[0][1];
  var razonCompra = registrosAprobados[0][3];
  var fechaRegistro = registrosAprobados[0][4];
  var justificacion = registrosAprobados[0][6];
  var prioridad = registrosAprobados[0][5];

  var date = new Date();
  var mes = date.getMonth() + 1;
  var anio = date.getFullYear();

  var nombreArchivo = `CAPEX_${registrosAprobados[0][0]}_${date.getDate()}-${date.getMonth() + 1}-${date.getFullYear()}.pdf`;

  // Reemplazar datos generales en el documento
  var body = doc.getBody();
  body.replaceText("{{solicitante}}", solicitante);
  body.replaceText("{{justificacion}}", justificacion);
  body.replaceText("{{prioridad}}", prioridad);
  body.replaceText("{{totalCompra}}", totalCompra.toFixed(2));
  body.replaceText("{{aprobador}}", aprobadoresEmail);
  body.replaceText("{{mes}}", mes);
  body.replaceText("{{anio}}", anio);
  body.replaceText("{{razonCompra}}", razonCompra.toUpperCase());

  // Crear un bloque de texto repetible
  var productosPlaceholder = "{{#productos}}";
  var startIndex = body.getText().indexOf(productosPlaceholder);
  var endIndex = body.getText().indexOf("{{/productos}}") + "{{/productos}}".length;
  var blockText = body.getText().substring(startIndex, endIndex);
  var productTemplate = blockText.replace(productosPlaceholder, "").replace("{{/productos}}", "");


  // Crear un bloque de texto repetible
  var productContent = "";
  var descripAllProducts = "";
  for (var i = 0; i < registrosAprobados.length; i++) {
    var producto = registrosAprobados[i];
    var cantidad = producto[11];
    var centroDeCosto = producto[10];
    var equipo = producto[7];
    var marca = producto[8];
    var especificaciones = producto[9];
    var precio = producto[12];
    var subtotal = producto[13];

    var productEntry = + cantidad + "  " + equipo.toUpperCase() + marca.toUpperCase() + ((i + 1 < registrosAprobados.length) ? "," : "");

    // Formato del descripAllProductsEntry
    var descripAllProductsEntry = centroDeCosto.toUpperCase()+ "\n"+ "- " + cantidad + " " + equipo + " " + marca + "  " + especificaciones + "\n" +
      "    " + "Unit Price: US$: " + precio + "\n" +
      "    " + "Sub Total: US$: " + subtotal + "\n";

    // Añadir saltos de línea entre productos
    /*if (i + 1 < registrosAprobados.length) {
      descripAllProductsEntry += "\n\n";
    }*/

    productContent += productEntry;
    descripAllProducts += descripAllProductsEntry;
  }

  body.replaceText("{{PRODUCTOS}}", productContent);
  body.replaceText("{{DESCRIPTALLPRODUCTS}}", descripAllProducts);



  // Guardar cambios
  doc.saveAndClose();

  // Convertir el documento a PDF
  var pdf = copiaPlantilla.getAs(MimeType.PDF);

  // Crear el archivo PDF
  var archivoPdf = carpetaPdf.createFile(pdf);

  // Cambiar el nombre del archivo
  archivoPdf.setName(nombreArchivo);

  // Obtener el enlace del archivo PDF
  var linkPdf = archivoPdf.getUrl();
  console.log(" LINK DE PDF " + linkPdf);
  console.log(" ARCHIVO PDF " + archivoPdf);


  return { pdf: archivoPdf, link: linkPdf };

}

function generCapex() {
  //Toma de datos de la fila activa
  var colNames = 2;
  var colEmail = 3;
  var colrCompra = 4;
  var colFechaRegistro = 5;
  var colPrioridad = 6;
  var colCentroCosto = 7;
  var colEquipos = 8;
  var colMarca = 9;
  var colEspeficaciones = 10;
  var colCantidad = 11;
  var colPrecio = 12;
  var colSubTotal = 13;
  var colAprobadoPor = 14;
  var colFechaAprobacion = 15;

  //Identificaciones 
  var capexPlantillaId = "1f4gccifvzDY23aWDhEaeEHSNzEEgW-vjdVIQE00HsPk";
  var pdfId = "1T_vTy4BVj3ypQbes5jMm4yWYOaWmacF3";
  var tempId = "1PN_mQmT_RoOZJCRwhPd0xr8zkeNrNeIP";

  //Para poder realizar cambios en google docs
  var document = DocumentApp.openById(capexPlantillaId);
  var capexPlantilla = DriveApp.getFileById(capexPlantillaId);
  var carpetaPdf = DriveApp.getFolderById(pdfId);
  var carpetaTemp = DriveApp.getFolderById(tempId);
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("idex");

  //Variables para el documento 
  var filaActiva = hoja.getActiveRange().getRow();
  var names = hoja.getRange(filaActiva, colNames).getValue();
  var email = hoja.getRange(filaActiva, colEmail);
  var rCompra = hoja.getRange(filaActiva, colrCompra);
  var fechaRegistro = hoja.getRange(filaActiva, colFechaRegistro);
  var prioridad = hoja.getRange(filaActiva, colPrioridad);
  var centroCosto = hoja.getRange(filaActiva, colCentroCosto);
  var equipos = hoja.getRange(filaActiva, colEquipos).getValue();
  var marca = hoja.getRange(filaActiva, colMarca);
  var especificaciones = hoja.getRange(filaActiva, colEspeficaciones).getValue();
  var cantidad = hoja.getRange(filaActiva, colCantidad).getValue();
  var precio = hoja.getRange(filaActiva, colPrecio).getValue();
  var subtotal = hoja.getRange(filaActiva, colSubTotal).getValue();
  var aprobador = hoja.getRange(filaActiva, colAprobadoPor).getValue();
  var fechaAprobacion = hoja.getRange(filaActiva, colFechaAprobacion);

  //
  var copiaPlantilla = capexPlantilla.makeCopy(carpetaTemp);
  var copiaId = copiaPlantilla.getId();
  var doc = DocumentApp.openById(copiaId);

  //Reemplazar variables
  doc.getBody().replaceText("{{names}}", names);
  doc.getBody().replaceText("{{cantidad}}", cantidad);
  doc.getBody().replaceText("{{equipos}}", equipos);
  doc.getBody().replaceText("{{especificaciones}}", especificaciones);
  doc.getBody().replaceText("{{aprobador}}", aprobador);
  doc.getBody().replaceText("{{precio}}", precio);
  doc.getBody().replaceText("{{names}}", names);
  doc.getBody().replaceText("{{subtotal}}", subtotal);
}







