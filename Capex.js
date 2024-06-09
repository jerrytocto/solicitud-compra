function generateCapex(totalCompra, registrosAprobados, aprobadoresEmail, conFirma) {


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

  console.log("Registros en el archivo capex: "+ registrosAprobados)
  // Datos generales (asumiendo que los datos generales están en la primera fila)
  var solicitante = registrosAprobados[0][1];
  var razonCompra = registrosAprobados[0][3];
  var fechaRegistro = registrosAprobados[0][4];
  var justificacion = registrosAprobados[0][6];
  var prioridad = registrosAprobados[0][5];
  var solicitudId = registrosAprobados[0][0];
  var observaciones = registrosAprobados[0][14];

  //Obtener fecha de creación del capex
  var date = new Date();
  var dia = date.getDate();
  var mes = date.getMonth() + 1;
  var anio = date.getFullYear();

  

  // Reemplazar datos generales en el documento
  var body = doc.getBody();
  body.replaceText("{{solicitante}}", solicitante);
  body.replaceText("{{justificacion}}", justificacion);
  body.replaceText("{{prioridad}}", prioridad);
  body.replaceText("{{totalCompra}}", totalCompra.toFixed(2));
  body.replaceText("{{aprobador}}", aprobadoresEmail[0]);
  body.replaceText("{{solicitudId}}", solicitudId);
  body.replaceText("{{dia}}", dia);
  body.replaceText("{{mes}}", mes);
  body.replaceText("{{anio}}", anio);
  body.replaceText("{{razonCompra}}", razonCompra.toUpperCase());
  body.replaceText("{{observaciones}}", observaciones);

  if (conFirma) {
    console.log("Se tiene que firmar")
    body.replaceText("{{gerenteGeneral}}", aprobadoresEmail[1]);
    body.replaceText("{{date}}", date);
    var nombreArchivo = `CAPEX_FIRM_${solicitudId}_${dia}-${mes}-${anio}.pdf`;
  } else{
    var nombreArchivo = `CAPEX_${solicitudId}_${dia}-${mes}-${anio}.pdf`;
  }

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
    var descripAllProductsEntry = centroDeCosto.toUpperCase() + "\n" + "- " + cantidad + " " + equipo + " " + marca + "  " + especificaciones + "\n" +
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


  return { pdf: archivoPdf, link: linkPdf };

}






