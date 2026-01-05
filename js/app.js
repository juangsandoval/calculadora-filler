(async () => {

const logEl = document.getElementById("log");

window.calcularYGestionar = async function () {
    const fNot = new Date(fechaNotificacion.value);
    const dias = parseInt(diasTermino.value, 10);
    const fPres = new Date(fechaPresentacion.value);

    if (!fNot || !fPres || isNaN(dias)) {
        log("Datos incompletos");
        return;
    }

    const vencimiento = new Date(fNot);
    vencimiento.setDate(vencimiento.getDate() + dias);

    let clasificacion;
    if (fPres < vencimiento) clasificacion = "pretermino";
    else if (fPres.getTime() === vencimiento.getTime()) clasificacion = "entiempo";
    else clasificacion = "extemporaneo";

    log(`Clasificación: ${clasificacion.toUpperCase()}`);

    const plantilla = `plantillas/${clasificacion}.docx`;
    const zip = await cargarPlantilla(plantilla);

    const contexto = {
        fecha_notificacion: formatDate(fNot),
        fecha_vencimiento: formatDate(vencimiento),
        fecha_presentacion: formatDate(fPres),
        clasificacion: clasificacion.toUpperCase()
    };

    const blob = generarDOCX(zip, contexto);
    saveAs(blob, `resultado_${clasificacion}.docx`);
};

async function cargarPlantilla(ruta) {
    const res = await fetch(ruta);
    if (!res.ok) throw new Error("No se pudo cargar plantilla");
    const buffer = await res.arrayBuffer();
    return new PizZip(buffer);
}

/* ================= MOTOR PROPIO ================= */

function generarDOCX(zip, data) {
    const files = Object.keys(zip.files)
        .filter(n => n.startsWith("word/") && n.endsWith(".xml"));

    files.forEach(name => {
        const xml = zip.file(name).asText();
        zip.file(name, reemplazar(xml, data));
    });

    return zip.generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });
}

function reemplazar(xml, datos) {
  let resultado = xml;

  Object.keys(datos).forEach(clave => {
    const valor = datos[clave];
    const regex = new RegExp(`{{\\s*${clave}\\s*}}`, "g");
    resultado = resultado.replace(regex, valor);
  });

  return resultado;
}

function aplicar(arr, start, len, val) {
    let acc = 0;
    for (let i = 0; i < arr.length; i++) {
        if (start >= acc && start < acc + arr[i].length) {
            const end = start + len;
            arr[i] = arr[i].slice(0, start - acc) + val;
            let j = i;
            let pos = acc + arr[i].length;
            while (pos < end && ++j < arr.length) {
                pos += arr[j].length;
                arr[j] = "";
            }
            break;
        }
        acc += arr[i].length;
    }
}

/* ================= UTILIDADES ================= */

function formatDate(d) {
    return d.toLocaleDateString("es-CO", { day:"numeric", month:"long", year:"numeric" });
}

function escapeXml(s) {
    return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}

function log(m) { logEl.textContent += m + "\n"; }

})();

