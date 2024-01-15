const fs = require('fs');
const pdfParse = require('pdf-parse');
const ExcelJS = require('exceljs');

// Patrones de Regex para extracción de datos
const regexInformacionCompleta = /(\d+)(\d{2}\/\d{2}\/\d{4} \d{2}:\d{2})([A-Za-z0-9\s]+[A-Za-zÀ-ÿ\u00f1\u00d1]+\s+)/;
const regexDestino = /\b[0-9]{4}(\s[A-Za-z]+)+\b/;
const regexDestinoReal = /ESP\d\d\d\d\d ([A-Za-z]+( [A-Za-z]+)+)/;
const regexNumeroDeViaje = /\b\d{4}[A-Za-z]{4}\d{4}\b/;
const regexNumeroDePedido = /\d\d\d\d\d\d\d\d\/\d/;
const regexMatricula = /^[0-9]{1,4}(?!.*(LL|CH))[BCDFGHJKLMNPRSTVWXYZ]{3}/;

// Función para leer el PDF y extraer el texto
async function readPdf(filePath) {
    const dataBuffer = fs.readFileSync(filePath);
    return await pdfParse(dataBuffer);
}

// Función para procesar el texto del PDF y convertirlo en una estructura de datos
function processData(text) {
    const lines = text.split('\n');
    const camiones = [];
    let currentCamion = {};
    lines.forEach(line => {
        let matchResult;

        if ((matchResult = line.match(regexInformacionCompleta))) {
            currentCamion.pallets = matchResult[1];
            currentCamion.fechaRecogida = matchResult[2];
            currentCamion.delegacion = matchResult[3].trim();
        } else if ((matchResult = line.match(regexDestino)) || (matchResult = line.match(regexDestinoReal))) {
            if (currentCamion.destino) camiones.push(currentCamion);
            currentCamion = { destino: matchResult[0] };
        } else if ((matchResult = line.match(regexNumeroDePedido))) {
            currentCamion.numeroDePedido = matchResult[0];
        } else if ((matchResult = line.match(regexNumeroDeViaje))) {
            currentCamion.numeroDeViaje = matchResult[0];
        } else if ((matchResult = line.match(regexMatricula))) {
            currentCamion.matricula = matchResult[0];
        }
    });

    if (Object.keys(currentCamion).length !== 0) camiones.push(currentCamion);

    return camiones;
}

// Función para crear un archivo Excel con los datos de los camiones
async function createExcelFile(camiones, outputPath) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Camiones');

    // Añadir encabezados de columna
    sheet.columns = [
        { header: 'Destino', key: 'destino' },
        { header: 'Número de Pedido', key: 'numeroDePedido' },
        { header: 'Número de Viaje', key: 'numeroDeViaje' },
        { header: 'Matrícula', key: 'matricula' },
        { header: 'Fecha Recogida', key: 'fechaRecogida', style: { numFmt: 'dd/mm/yyyy hh:mm' } },
        { header: 'Pallets', key: 'pallets' },
        { header: 'Delegación', key: 'delegacion' },
        // Agregar más columnas si es necesario
    ];

    // Añadir datos a las filas
    camiones.forEach(camion => {
        sheet.addRow(camion);
    });

    // Guardar el archivo Excel
    await workbook.xlsx.writeFile(outputPath);
}

// Ruta al archivo PDF
const pdfPath = 'file.pdf';

// Procesar PDF y crear archivo Excel
readPdf(pdfPath)
    .then(data => {
        const camiones = processData(data.text);
        return createExcelFile(camiones, 'camiones.xlsx');
    })
    .then(() => {
        console.log('Archivo Excel creado con éxito.');
    })
    .catch(err => {
        console.error('Error:', err);
    });





