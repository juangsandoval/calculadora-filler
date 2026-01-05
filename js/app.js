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
    console.log("Contexto enviado al DOCX:", contexto);
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
/* ================= FUNCIÓN DE REEMPLAZAR ================= */
function reemplazar(xml, data) {
    const tRegex = /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/g;

    // 1. Extraer todos los runs <w:t>
    const runs = [];
    let match;
    while ((match = tRegex.exec(xml)) !== null) {
        runs.push({
            start: match.index,
            full: match[0],
            open: match[1],
            text: match[2],
            close: match[3]
        });
    }
    if (runs.length === 0) return xml;

    // 2. Unir todos los textos
    const texts = runs.map(r => r.text);
    const combined = texts.join("");

    // 3. Buscar placeholders {{ }}
    const regex = /{{\s*([\w_]+)\s*}}/g;
    let m;

    // Para no romper índices, reemplazamos de atrás hacia adelante
    const replacements = [];
    while ((m = regex.exec(combined))) {
        replacements.push({
            key: m[1],
            start: m.index,
            end: m.index + m[0].length
        });
    }

    if (replacements.length === 0) return xml;

    // 4. Aplicar reemplazos sobre el array de textos
    for (let i = replacements.length - 1; i >= 0; i--) {
        const { key, start, end } = replacements[i];
        const value = escapeXml(data[key] ?? "");

        let acc = 0;
        let startRun = -1;
        let endRun = -1;
        let startOffset = 0;
        let endOffset = 0;

        // Encontrar runs involucrados
        for (let r = 0; r < texts.length; r++) {
            const len = texts[r].length;
            if (startRun === -1 && start >= acc && start < acc + len) {
                startRun = r;
                startOffset = start - acc;
            }
            if (endRun === -1 && end > acc && end <= acc + len) {
                endRun = r;
                endOffset = end - acc;
            }
            acc += len;
        }

        if (startRun === -1 || endRun === -1) continue;

        // Reemplazar
        texts[startRun] =
            texts[startRun].slice(0, startOffset) +
            value +
            texts[endRun].slice(endOffset);

        // Vaciar runs intermedios
        for (let r = startRun + 1; r <= endRun; r++) {
            texts[r] = "";
        }
    }

    // 5. Reconstruir XML
    let output = "";
    let cursor = 0;

    for (let i = 0; i < runs.length; i++) {
        const r = runs[i];
        output += xml.slice(cursor, r.start);
        output += r.open + texts[i] + r.close;
        cursor = r.start + r.full.length;
    }

    output += xml.slice(cursor);
    return output;
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




