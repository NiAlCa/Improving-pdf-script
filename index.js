const fs = require('fs');
const pdfParse = require('pdf-parse');
const ExcelJS = require('exceljs');
const path = require('path');

const folderPath = process.cwd();

//Patrones de Regex
const regexInformacionCompleta = /(\d+)(\d{2}\/\d{2}\/\d{4} \d{2}:\d{2})([A-Za-z0-9\s]+[A-Za-zÀ-ÿ\u00f1\u00d1]+\s+)/;
const regexDestino = /\b[0-9]{4}(\s[A-Za-z]+)+\b/;
const regexDestinoReal = /ESP\d\d\d\d\d ([A-Za-z]+( [A-Za-z]+)+)/;
const regexPedidoViaje = /(\d{8}\/\d)([a-zA-Z0-9]+)/; 
const regexMatricula = /^[0-9]{1,4}(?!.*(LL|CH))[BCDFGHJKLMNPRSTVWXYZ]{3}/;
const regexAlmacen1 = /8412353000187/;
const regexAlmacen2 = /8412353000170/;
const regexAlmacen3 = /8412353000033/;

// Función para leer el PDF
async function readPdf(filePath) {
    const dataBuffer = fs.readFileSync(filePath);
    return await pdfParse(dataBuffer);
}

// Función para detectar el almacén 
function detectAlmacen(text) {
    if (regexAlmacen1.test(text)) return '8412353000187';
    if (regexAlmacen2.test(text)) return '8412353000170';
    if (regexAlmacen3.test(text)) return '8412353000033';
    return ''; 
}

// Función para procesar el texto del PDF y convertirlo en una estructura de datos
function processData(text) {
    const lines = text.split('\n');
    const camiones = [];
    let currentCamion = {};

    // Detectar el almacén para este PDF
    const almacenDetectado = detectAlmacen(text);

    lines.forEach(line => {
        let matchResult;

        if ((matchResult = line.match(regexInformacionCompleta))) {
            currentCamion.pallets = matchResult[1];
            currentCamion.fechaRecogida = matchResult[2];
            currentCamion.delegacion = matchResult[3].trim();
        } else if ((matchResult = line.match(regexDestino)) || (matchResult = line.match(regexDestinoReal))) {
            if (currentCamion.destino) {
                currentCamion.almacen = almacenDetectado; 
                camiones.push(currentCamion);
            }
            currentCamion = { destino: matchResult[0] };
        } else if ((matchResult = line.match(regexPedidoViaje))) {
            currentCamion.numeroDePedido = matchResult[1];
            currentCamion.numeroDeViaje = matchResult[2];
        } else if ((matchResult = line.match(regexMatricula))) {
            currentCamion.matricula = matchResult[0];
        }
    });

    if (Object.keys(currentCamion).length !== 0) {
        currentCamion.almacen = almacenDetectado; 
        camiones.push(currentCamion);
    }

    return camiones;
}

function tieneCoincidencias(numeroDeViaje, camiones) {
    return camiones.filter(camion => camion.numeroDeViaje === numeroDeViaje).length > 1;
}
const colors = [
    'FFFF00', // Amarillo
    'FF0000', // Rojo
    '00FF00', // Verde
    '0000FF', // Azul
    'FF00FF', // Magenta
    '00FFFF', // Cian
    'FF8000', // Naranja
    '8000FF', // Morado
    'A52A2A', // Marrón
    '808080', // Gris
    '000000', // Negro
    'FFFFFF', // Blanco
    'FFC0CB', // Rosa
    'FA8072', // Salmón
    'FFE5B4', // Melocotón
    '808000', // Verde oliva
    '000080', // Azul marino
    'C8A2C8', // Lila
    'FFD700', // Amarillo vibrante
    'DC143C', // Rojo vibrante
    '00FF7F', // Verde vibrante
    '007FFF', // Azul vibrante
    'FF0080', // Magenta vibrante
    '00FFFF', // Cian vibrante
  


];

// Función para asignar un color único a cada número de pedido
function getColorForPedido(numeroDeViaje, pedidoColorsMap) {
    if (!pedidoColorsMap.has(numeroDeViaje)) {
        // Asignar un nuevo color de la lista, rotando si es necesario
        const color = colors[pedidoColorsMap.size % colors.length];
        pedidoColorsMap.set(numeroDeViaje, color);
    }
    return pedidoColorsMap.get(numeroDeViaje);
}

async function createExcelFile(camiones, outputPath) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Camiones');
    const pedidoColorsMap = new Map(); // Mapa para asociar número de pedido con color

    // Añadir encabezados de columna
    sheet.columns = [
        { header: 'Destino', key: 'destino' },
        { header: 'Número de Pedido', key: 'numeroDePedido' },
        { header: 'Número de Viaje', key: 'numeroDeViaje' },
        { header: 'Matrícula', key: 'matricula' },
        { header: 'Fecha Recogida', key: 'fechaRecogida', style: { numFmt: 'dd/mm/yyyy hh:mm' } },
        { header: 'Pallets', key: 'pallets' },
        { header: 'Delegación', key: 'delegacion' },
        { header: 'Almacén', key: 'almacen' },
    ];

    camiones.forEach(camion => {
        const row = sheet.addRow(camion);
        const color = getColorForPedido(camion.numeroDeViaje, pedidoColorsMap);
        row.eachCell(cell => {
            cell.style.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: color }
            };
        });
    });

    // Guardar el archivo Excel
    await workbook.xlsx.writeFile(outputPath);
    console.log('Archivo Excel creado con éxito.');
}

// Función para procesar todos los PDFs en una carpeta
async function processAllPdfsInFolder(folderPath) {
const files = fs.readdirSync(folderPath);
const pdfFiles = files.filter(file => path.extname(file).toLowerCase() === '.pdf');

let allCamiones = [];

for (let file of pdfFiles) {
    const filePath = path.join(folderPath, file);
    try {
        const data = await readPdf(filePath);
        const camiones = processData(data.text);
        allCamiones = allCamiones.concat(camiones);
    } catch (err) {
        console.error('Error procesando el archivo:', file, err);
    }
}
return allCamiones;



}



// Procesar todos los PDFs en la carpeta y crear archivo Excel
processAllPdfsInFolder(folderPath)
.then(allCamiones => {
return createExcelFile(allCamiones, path.join(folderPath, 'camiones.xlsx'));
})
.catch(err => {
console.error('Error:', err);
});

