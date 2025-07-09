# pdf_makers
Repositorio de codigos funcionales (2025) para automatización o generación manual de PDF con hojas de google spread sheets y App Scripts


//Para base de talento humano permisos generados a través de GTR

function generarPDF() {
  const hojaNombre = "adm_tthh_permisos";
  const fila = 335; // Fila específica a procesar
  const plantillaId = "seliminaporcuestionesdeeguridad";
  const carpetaId = "seliminaporcuestionesdeeguridad";

  try {
    // Obtener datos de la hoja de cálculo
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(hojaNombre);
    if (!hoja) throw new Error("La hoja de cálculo no existe.");

    const datos = hoja.getDataRange().getValues(); // Obtener datos en su formato original
    if (fila > datos.length || fila <= 0) throw new Error(`La fila ${fila} está fuera del rango.`);

    const filaDatos = datos[fila - 1]; // Ajuste de índice

    // Validar que la fila tiene datos
    if (!filaDatos || filaDatos.every(celda => !celda || celda.toString().trim() === "")) {
      throw new Error(`La fila ${fila} está vacía o no contiene datos.`);
    }

    Logger.log("Datos obtenidos de la fila: " + JSON.stringify(filaDatos));

    // Crear copia temporal de la plantilla
    const nombreArchivo = `Permiso_${filaDatos[3] || "Desconocido"}_${fila}`;
    const copia = DriveApp.getFileById(plantillaId)
      .makeCopy(nombreArchivo, DriveApp.getFolderById(carpetaId));

    // Abrir el documento
    const doc = DocumentApp.openById(copia.getId());
    const cuerpo = doc.getBody();
    const encabezado = doc.getHeader();
    const piePagina = doc.getFooter();
    const tablas = cuerpo.getTables(); // Obtener todas las tablas del documento

    // Funciones para obtener valores correctos
    const obtenerValor = (valor) => (valor && valor.toString().trim() !== "" ? valor : "N/A");
    const formatoFecha = (valor) => {
      if (!valor || valor === "N/A") return valor;
      return Utilities.formatDate(new Date(valor), Session.getScriptTimeZone(), "dd/MM/yyyy");
    };
    const formatoHora = (valor) => {
      if (!valor || valor === "N/A") return valor;
      return Utilities.formatDate(new Date(valor), Session.getScriptTimeZone(), "HH:mm");
    };

    const reemplazos = [
      ["<<[FUNCIONARIO]>>", obtenerValor(filaDatos[3])],
      ["<<[FECHA]>>", formatoFecha(filaDatos[1])],
      ["<<[CEDULA]>>", obtenerValor(filaDatos[4])],
      ["<<[AREA DE TRABAJO]>>", obtenerValor(filaDatos[5])],
      ["<<[REGIMEN]>>", obtenerValor(filaDatos[7])],
      ["<<[TIPO DE PERMISO]>>", obtenerValor(filaDatos[8])],
      ["<<[DESDE FECHA]>>", formatoFecha(filaDatos[9])],
      ["<<[DESDE HORA]>>", formatoHora(filaDatos[10])],
      ["<<[HASTA FECHA]>>", formatoFecha(filaDatos[11])],
      ["<<[HASTA HORA]>>", formatoHora(filaDatos[12])],
      ["<<[MOTIVO DEL PERMISO]>>", obtenerValor(filaDatos[13])]
    ];

    // Reemplazo en cuerpo, encabezado, pie de página y tablas
    reemplazos.forEach(([marcador, valor]) => {
      Logger.log(`Reemplazando ${marcador} con ${valor}`);
      cuerpo.replaceText(marcador, valor);
      if (encabezado) encabezado.replaceText(marcador, valor);
      if (piePagina) piePagina.replaceText(marcador, valor);

      // Reemplazo en todas las tablas
      tablas.forEach(tabla => {
        for (let i = 0; i < tabla.getNumRows(); i++) {
          for (let j = 0; j < tabla.getRow(i).getNumCells(); j++) {
            let celda = tabla.getRow(i).getCell(j);
            if (celda.getText().includes(marcador)) {
              Logger.log(`Reemplazando ${marcador} en una tabla`);
              celda.setText(celda.getText().replace(marcador, valor));
            }
          }
        }
      });
    });

    // Guardar cambios en el documento antes de la conversión a PDF
    doc.saveAndClose();

    // **Esperar más tiempo para asegurar que los cambios se reflejen antes de convertir a PDF**
    Utilities.sleep(7000);

    // Obtener el archivo actualizado antes de convertirlo en PDF
    const archivoActualizado = DriveApp.getFileById(copia.getId());
    const pdfBlob = archivoActualizado.getAs("application/pdf").setName(`${nombreArchivo}.pdf`);
    DriveApp.getFolderById(carpetaId).createFile(pdfBlob);

    // Eliminar documento temporal
    archivoActualizado.setTrashed(true);

    console.log(`✅ PDF generado correctamente: ${nombreArchivo}`);

  } catch (error) {
    console.error(`❌ Error en la generación del PDF para la fila ${fila}:`, error.message);
  }
}



//Codigo en hoja de solicitudes de información publica para generar pdf y enviar a correos la solicitud.

function generarPDFyEnviarCorreo(fila) {
  const hojaNombre = "SolicitudesInfo"; // Cambia al nombre de la nueva hoja
  const plantillaId = "1OORVO2c_4ZcAj5OgNKp2RiEVx74xUdQrAHyzTs3Fo-I"; // ID de la plantilla
  const carpetaId = "1__RDFSm3fAtYgfKz7jM8gt3koJ9KbMYT"; // ID de la carpeta

  try {
    // Obtener datos de la hoja de cálculo
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(hojaNombre);
    if (!hoja) throw new Error("La hoja de cálculo no existe.");

    const datos = hoja.getDataRange().getValues(); // Obtener todos los datos
    if (fila > datos.length || fila <= 0) throw new Error(`La fila ${fila} está fuera del rango.`);

    const filaDatos = datos[fila - 1]; // Ajuste de índice
    if (!filaDatos || filaDatos.every(celda => !celda || celda.toString().trim() === "")) {
      throw new Error(`La fila ${fila} está vacía o no contiene datos.`);
    }

    Logger.log("Datos obtenidos de la fila: " + JSON.stringify(filaDatos));

    // Verificar si la columna "Generado" (columna N, índice 13) ya está marcada
    const columnaGenerado = filaDatos[13]; // Columna N
    if (columnaGenerado && columnaGenerado.toString().trim().toUpperCase() === "SI") {
      console.log(`⚠️ El PDF para la fila ${fila} ya fue generado. Saltando...`);
      return;
    }

    // Crear copia temporal de la plantilla
    const nombreArchivo = `Solicitud_${filaDatos[3]}_${fila}`; // Usando Nombres como referencia
    const copia = DriveApp.getFileById(plantillaId)
      .makeCopy(nombreArchivo, DriveApp.getFolderById(carpetaId));

    // Abrir el documento
    const doc = DocumentApp.openById(copia.getId());
    const cuerpo = doc.getBody();

    // Funciones para obtener valores correctos
    const obtenerValor = (valor) => (valor && valor.toString().trim() !== "" ? valor : "N/A");

    // Mapeo de marcadores según la nueva estructura
    const reemplazos = [
      ["<<Marca temporal>>", obtenerValor(filaDatos[0])], // Columna A (Marca temporal)
      ["<<Fecha>>", obtenerValor(filaDatos[1])],         // Columna B (Fecha)
      ["<<Ciudad>>", obtenerValor(filaDatos[2])],        // Columna C (Ciudad)
      ["<<Nombres>>", obtenerValor(filaDatos[3])],       // Columna D (Nombres)
      ["<<Apellidos>>", obtenerValor(filaDatos[4])],     // Columna E (Apellidos)
      ["<<Cédula>>", obtenerValor(filaDatos[5])],        // Columna F (Cédula)
      ["<<Dirección domiciliaria>>", obtenerValor(filaDatos[6])], // Columna G
      ["<<Teléfono (fijo o celular)>>", obtenerValor(filaDatos[7])], // Columna H
      ["<<Describa la petición>>", obtenerValor(filaDatos[8])], // Columna I
      ["<<Forma de recepción>>", obtenerValor(filaDatos[9])],   // Columna J
      ["<<Formato de entrega>>", obtenerValor(filaDatos[10])],  // Columna K
      ["<<Si es Formato digital>>", obtenerValor(filaDatos[11])], // Columna L
      ["<<Correo electrónico>>", obtenerValor(filaDatos[12])]   // Columna M
    ];

    // Reemplazo de marcadores en el cuerpo del documento
    reemplazos.forEach(([marcador, valor]) => {
      cuerpo.replaceText(marcador, valor);
    });

    // Guardar cambios en el documento
    doc.saveAndClose();

    // Esperar hasta que el archivo esté listo
    esperarArchivoListo(copia.getId());

    // Convertir a PDF
    const archivoActualizado = DriveApp.getFileById(copia.getId());
    const pdfBlob = archivoActualizado.getAs("application/pdf").setName(`${nombreArchivo}.pdf`);
    const archivoPDF = DriveApp.getFolderById(carpetaId).createFile(pdfBlob);

    // Eliminar documento temporal
    archivoActualizado.setTrashed(true);

    console.log(`✅ PDF generado correctamente: ${nombreArchivo}`);

    // Enviar correo electrónico
    enviarCorreoConAdjunto(
      filaDatos[12], // Correo electrónico (Columna M)
      filaDatos[8],  // Descripción de la petición (Columna I)
      archivoPDF.getUrl(),
      pdfBlob
    );

    // Actualizar la columna "Generado" (columna N, índice 13) con "SI"
    hoja.getRange(fila, 14).setValue("SI"); // Columna N
    console.log(`✅ Columna "Generado" actualizada para la fila ${fila}`);

  } catch (error) {
    console.error(`❌ Error en la generación del PDF para la fila ${fila}:`, error.message);
  }
}

// Función para enviar correo electrónico
function enviarCorreoConAdjunto(destinatario, motivo, urlPDF, adjunto) {
  const remitente = Session.getActiveUser().getEmail(); // Correo del usuario activo
  const copia1 = "informacion@gadconocoto.gob.ec"; // Primer correo en copia
  const copia2 = "tecnologia@gadconocoto.gob.ec"; // Segundo correo en copia

  const asunto = `Solicitud Generada: ${motivo}`;
  const cuerpoMensaje = `
    Estimado/a,

    Se ha generado una solicitud con la siguiente descripción:
    ${motivo}

    Puede descargar el documento de solicitud PDF aquí:
    ${urlPDF}

    Saludos,
    TICS
  `;

  MailApp.sendEmail({
    to: destinatario,
    cc: `${copia1}, ${copia2}`,
    subject: asunto,
    body: cuerpoMensaje,
    attachments: [adjunto]
  });

  console.log(`Correo enviado a ${destinatario} con copia a ${copia1} y ${copia2}`);
}

// Función para esperar hasta que el archivo esté listo
function esperarArchivoListo(fileId, intentos = 10, retraso = 1000) {
  for (let i = 0; i < intentos; i++) {
    const archivo = DriveApp.getFileById(fileId);
    if (archivo.getSize() > 0) return true;
    Utilities.sleep(retraso);
  }
  throw new Error("El archivo no pudo guardarse correctamente.");
}

// Menú personalizado para activar el trigger manual
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Generar PDF")
    .addItem("Generar PDF Manualmente", "mostrarDialogoSeleccionFila")
    .addToUi();
}

// Diálogo para seleccionar la fila
function mostrarDialogoSeleccionFila() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial; padding: 20px;">
      <h3>Generar PDF</h3>
      <p>Ingrese el número de fila (sin incluir los encabezados):</p>
      <input type="number" id="fila" min="2" placeholder="Ej. 2" style="width: 100%; padding: 8px;" />
      <br><br>
      <button 
        type="button" 
        onclick="enviarFila()" 
        style="padding: 10px 20px; background-color: #4CAF50; color: white; border: none; cursor: pointer;">
        Generar
      </button>
      <script>
        function enviarFila() {
          const fila = parseInt(document.getElementById("fila").value);
          if (!isNaN(fila) && fila > 1) {
            google.script.run
              .withSuccessHandler(() => google.script.host.close())
              .generarPDFyEnviarCorreo(fila);
          } else {
            alert("Por favor, ingrese un número de fila válido (mayor que 1).");
          }
        }
      </script>
    </div>
  `).setWidth(400).setHeight(250);

  SpreadsheetApp.getUi().showModalDialog(html, "Seleccionar Fila");
}



