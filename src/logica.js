const newWorkbook = XLSX.utils.book_new() // crear nuevo archivo Excel
const newSheet =XLSX.utils.aoa_to_sheet([]) // crear nueva hoja Excel


// Obtener una referencia al elemento de entrada de archivo
const fileLoadDetection = document.getElementById('fileLoadDetection');

fileLoadDetection.addEventListener('change', (e) => { // evento para detectar la carga del archivo
    const file = e.target.files[0]; // define variable con el archivo cargado

    if (file) {
        const reader = new FileReader(); // si file no es un error, se crea una plantilla de FileReader con sus metodos
        reader.readAsBinaryString(file); // leemos reader con el metodo readAsBinaryString de FileReader
        
        reader.onload = (e) => { // usamos el metodo onload de FileReader para detectar cuando el archivo es terminado de leer
            const data = e.target.result;  // guardamos en data la vinculacion a e.target.result que es la info en binario del excel.
            const workbook = XLSX.read(data, { type: 'binary' }); // guardamos en workbook el array resultado del metodo read de Sheet JS donde leemos data

                    
            const sheetName = workbook.SheetNames[0] // establezco un nombre o vinculo directo a la primer hoja del array/excel
            const sheet = workbook.Sheets[sheetName] // workbook.Sheets es un objeto que contiene representaciones de todas las hojas del excel y accedemos directo a la que nos vinculamos en el paso anterior.

            const jsonData = XLSX.utils.sheet_to_json(sheet, {header: 'A'}); // convierte los datos del Excel del objeto sheet en un array de objetos que van a representar las filas y columnas del Excel.
                                                                            // cada objeto es una fila, cada propiedad una columna y cada valor un valor.
            const jsonDataWithoyFirstRow = jsonData.slice(1)  // Crea copia de jsonData menos la primer fila para que no itere los encabezados
                                                                                
            const arrayDeArrays = [];
            const columnHeaders = ['Local','Item', 'Valor','Fecha',] // crea encabezado Excel
            arrayDeArrays.unshift(columnHeaders); // SUBE EL TITULO COMO PRIMERA FILA DEL ECEL PARA EL ENCABEZADO
                
            for (const fila of jsonDataWithoyFirstRow) { // recorre array (cada objeto es una fila)
                
                for (const columna in fila) { // recorre cada columna de cada fila (cada fila es un objeto)
                
                    if (columna !== 'A' && columna !== 'B') { // condicion para evitar que itere sobre las dos primeras columnas
                    const valor = fila[columna]; // asigna a valor la vinculacion del dato de cada celda iterada
                    
                        if (valor != 0 && valor != null && valor != '' && valor != '0') { // condicion para que solo se itere sobre valores distintos de cero
                        // Verifica si el valor es distinto de una cadena vacía, nulo o 0
                    
                        // Crea una nueva fila para cada valor
                        const nuevaFila = [];
            
                        // Agrega el valor de la celda A (en la primera columna)
                        nuevaFila.push(fila['A']);
            
                        // Agrega el valor de la celda B (en la segunda columna)
                        nuevaFila.push(fila['B']);
            
                        // Agrega el valor actual (en la tercera columna)
                        nuevaFila.push(valor);
                        
                        // Agrega el valor de fecha (correspondiente a la primera fila) que coincida con el valor de columna
                        nuevaFila.push(jsonData[0][columna]);

                        // Agrega la nueva fila al arrayDeArrays
                        arrayDeArrays.push(nuevaFila);                               
                        }
                    }
                }
            }

                /* DESCARGA DE ARCHIVOS */
                // Convertir arrayDeArrays en un formato CSV
                const csvContent = arrayDeArrays.map(row => row.join(',')).join('\n');

                // Crear un Blob a partir del contenido CSV
                const blob = new Blob([csvContent], { type: 'text/csv' });

                // Crear una URL para el Blob y generar un enlace de descarga
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'ConsolidadoVentasToteat-Output.csv'; // Nombre del archivo CSV
                document.body.appendChild(a);

                // Simular un clic en el enlace para iniciar la descarga
                a.click();

                // Liberar la URL del objeto Blob cuando ya no sea necesario
                window.URL.revokeObjectURL(url);

                // Quitar el enlace del DOM
                document.body.removeChild(a);
        };
    }
 
});





/* AJUSTES */

// Ver si puedo lograr que se baje directo en XLSX en vez del .CSV)
// Reiniciar input para que el script se pueda volver a usar sin actualizar.
// Pasarme las sentencias Sheet JS a notion









