function generarDocumento() {

  const fechaInicio = document.getElementById("fechaInicio").value;
  const fechaLimite = document.getElementById("fechaLimite").value;
  const fechaPresentacion = document.getElementById("fechaPresentacion").value;

  if (!fechaInicio || !fechaLimite || !fechaPresentacion) {
    alert("Por favor complete todas las fechas");
    return;
  }

  let estado = "";
  let conclusion = "";
  let plantilla = "";

  if (fechaPresentacion < fechaInicio) {
    estado = "PRETEMPORE";
    conclusion = "El escrito fue presentado de manera anticipada (pretempore).";
    plantilla = "plantillas/pretempore.docx";
  } 
  else if (fechaPresentacion <= fechaLimite) {
    estado = "EN TIEMPO";
    conclusion = "El escrito fue presentado dentro del término legal.";
    plantilla = "plantillas/pretempore.docx";
  } 
  else {
    estado = "EXTEMPORÁNEO";
    conclusion = "El escrito fue presentado por fuera del término legal.";
    plantilla = "plantillas/pretempore.docx";
  }

  cargarPlantillaYGenerar(plantilla, {
    fecha_inicio: fechaInicio,
    fecha_limite: fechaLimite,
    fecha_presentacion: fechaPresentacion,
    conclusion: conclusion
  }, estado);
}

function cargarPlantillaYGenerar(rutaPlantilla, datos, estado) {

  fetch(rutaPlantilla)
    .then(response => response.arrayBuffer())
    .then(buffer => {

      const zip = new PizZip(buffer);

      const doc = new window.docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true
      });

      doc.setData(datos);

      try {
        doc.render();
      } catch (error) {
        console.error(error);
        alert("Error al generar el documento Word");
        return;
      }

      const blob = doc.getZip().generate({
        type: "blob",
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      });

      saveAs(blob, `Resultado_${estado}.docx`);
    });
}
