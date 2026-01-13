document.getElementById('excel_graph').addEventListener('change', async function(event) {
    const file = event.target.files[0];
    
    if (!file) return;
    
    try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        const sheetSelect = document.getElementById('sheet_select');
        sheetSelect.innerHTML = '';
        
        workbook.SheetNames.forEach(sheetName => {
            const option = document.createElement('option');
            option.value = sheetName;
            option.textContent = sheetName;
            sheetSelect.appendChild(option);
        });
    } catch (error) {
        console.error('Error al leer el archivo Excel:', error);
    }
});

let sheetName = '';

document.getElementById('sheet_select').addEventListener('change', async function() {
    const file = document.getElementById('excel_graph').files[0];
    sheetName = this.value;
    
    if (!file || !sheetName) return;
    
    try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheet = workbook.Sheets[sheetName];
        const allData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        let lastRow = 0;
        allData.forEach((row, index) => {
            if (row[2] === 'TOTAL') lastRow = index;
        });
        
        const dataRows = allData.slice(2, lastRow);
        const tableData = [["Category", "Total Qty", "Total Amount", "Percentage", "Sysco Price", "Manufacturer", "Rank"]];
        
        const usdFormatter = new Intl.NumberFormat('en-US', {
            style: 'currency',
            currency: 'USD',
        });
        
        
        dataRows.forEach((row, index) => {
            tableData.push([
                row[2].length > 20 ? row[2].replace(/\s+/g, '\n') : row[2],
                Number(row[119]).toLocaleString('en-US'),
                usdFormatter.format(row[120]),
                (Math.round(row[121] * 1000) / 10) + '%',
                usdFormatter.format(row[122]),
                usdFormatter.format(row[123]),
                row[0]
            ]);
        });

        
        // displayTable(tableData);
        displayChart(tableData);
    } catch (error) {
        console.error('Error al procesar la hoja:', error);
    }
});

//const colors = ["#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#EAD1DC","#C9DAF8","#EAD1DC","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#EAD1DC","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#EAD1DC","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#C9DAF8","#FFF2CC","#EAD1DC","#EAD1DC","#C9DAF8","#C9DAF8","#FFF2CC"];

const colors = ["rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(234, 209, 220, 0.8)","rgba(201, 218, 248, 0.8)","rgba(234, 209, 220, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(234, 209, 220, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(234, 209, 220, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(255, 242, 204, 0.8)","rgba(234, 209, 220, 0.8)","rgba(234, 209, 220, 0.8)","rgba(201, 218, 248, 0.8)","rgba(201, 218, 248, 0.8)","rgba(255, 242, 204, 0.8)"];

function displayTable(data) {
    const container = document.getElementById('table_container');
    let html = '<table border="1" width="7108" style="table-layout: fixed"><tr>';
    
    for (let i = 0; i < 7; i++) {
        for(let j = 0; j < data.length; j++) {
            if (j == 0) {
                html += '<td width="56"><b>' + data[j][i] + '</b></td>';
            } else {
                html += '<td width="49" style="text-align:center">' + data[j][i] + '</td>';
            }
        }
        html += '</tr>';
    }
    
    html += '</table>';
    container.innerHTML += html;
}

function displayChart(data) {
    const chartData = data.slice(1).map(row => ({
        category: row[0],
        quantity: Number(row[1].replace(/,/g, ''))
    }));

    const ctx = document.createElement('canvas');
    ctx.width = 7050;
    ctx.height = 700;
    ctx.style.marginLeft = '57px';
    document.getElementById('graph_container').appendChild(ctx);

    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: chartData.map(item => item.category),
            datasets: [{
                label: 'Total Qty',
                data: chartData.map(item => item.quantity),
                backgroundColor: colors,
                borderColor: colors,
                borderWidth: 1
            }]
        },
        options: {
            responsive: false,
            plugins: {
                title: {
                    display: true,
                    text: 'GEN TOTAL QUANTITY (ALL STORES) ' + sheetName + ' 2026'
                },
                datalabels: {
                    anchor: 'end',
                    align: 'end',
                    formatter: (value) => value,
                    color: '#000'
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true, // Must be true to show the title
                        text: 'Total Qty (LBS)',
                        font: {
                            size: 20, // Set the font size here
                            weight: 'bold'
                        },
                        padding: { top: 10, bottom: 10 } // Optional padding
                    }
                },
                x : {
                    display: false,
                    ticks: {
                        minRotation: 0,
                        maxRotation: 0,
                        autoSkip: false // Set to false to show all labels and prevent automatic hiding/skipping
                    }
                }
            }
        },
        plugins: [ChartDataLabels]
    });

    displayTable(data);
}

async function downloadAsImage() {
    const graphContainer = document.getElementById('graph_container');
    const tableContainer = document.getElementById('table_container');
    
    if (!graphContainer.innerHTML || !tableContainer.innerHTML) {
        alert('Por favor, carga un archivo y selecciona una hoja primero');
        return;
    }
    
    try {
        // Obtener el canvas del gráfico y convertirlo a base64
        const chartCanvas = graphContainer.querySelector('canvas');
        const chartImageBase64 = chartCanvas.toDataURL('image/png');
        
        // Crear un contenedor temporal que combine gráfico y tabla
        const tempContainer = document.createElement('div');
        tempContainer.style.position = 'absolute';
        tempContainer.style.left = '-9999px';
        tempContainer.style.backgroundColor = 'white';
        tempContainer.style.padding = '20px';
        tempContainer.style.width = '7200px';
        
        // Crear imagen del gráfico
        const graphImg = document.createElement('img');
        graphImg.src = chartImageBase64;
        graphImg.style.width = '7050px';
        graphImg.style.marginBottom = '0px';
        graphImg.style.marginLeft = '57px';
        graphImg.style.marginRight = '20px';
        
        // Clonar la tabla
        const tableClone = tableContainer.cloneNode(true);
        
        tempContainer.appendChild(graphImg);
        tempContainer.appendChild(tableClone);
        document.body.appendChild(tempContainer);
        
        // Convertir a imagen
        const canvas = await html2canvas(tempContainer, {
            scale: 2,
            backgroundColor: '#ffffff',
            logging: false
        });
        
        // Limpiar el contenedor temporal
        document.body.removeChild(tempContainer);
        
        // Descargar la imagen
        const link = document.createElement('a');
        link.href = canvas.toDataURL('image/png');
        link.download = `grafica_${new Date().getTime()}.png`;
        link.click();
        
    } catch (error) {
        console.error('Error al descargar la imagen:', error);
        alert('Error al descargar la imagen. Intenta de nuevo.');
    }
}

// Agregar el evento al botón de descarga
document.addEventListener('DOMContentLoaded', () => {
    const downloadBtn = document.getElementById('download_btn');
    if (downloadBtn) {
        downloadBtn.addEventListener('click', downloadAsImage);
    }
});

