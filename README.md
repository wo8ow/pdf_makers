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




