document.addEventListener('DOMContentLoaded', () => {
    const calcularBtn = document.getElementById('calcularBtn');
    const resultadosDiv = document.getElementById('resultados');
    const nivelRiesgoSpan = document.getElementById('nivelRiesgo');
    const efectosRiesgoP = document.getElementById('efectosRiesgo');
    const medidasRiesgoUl = document.getElementById('medidasRiesgo');
    const descargarWORDBtn = document.getElementById('descargarDOC');

    let datosRubrica = null;

    // 1. Cargar la rúbrica del JSON
    async function cargarRubrica() {
        try {
            const response = await fetch('rubrica.json');
            datosRubrica = await response.json();
        } catch (error) {
            console.error('Error al cargar la rúbrica:', error);
            alert('Error: No se pudo cargar el archivo de datos (rubrica.json).');
        }
    }

    // 2. Función principal de cálculo
    function calcularRiesgo() {
        if (!datosRubrica) {
            alert('Datos de cálculo no cargados. Intente recargar la página.');
            return;
        }

        // Obtener valores de entrada
        const temp = parseFloat(document.getElementById('temperatura').value);
        const hum = parseFloat(document.getElementById('humedad').value);

        if (isNaN(temp) || isNaN(hum) || temp <= 0 || hum <= 0) {
            alert('Por favor, ingrese valores válidos para Temperatura y Humedad.');
            return;
        }

        const { temperaturas, humedad_relativa, matriz_riesgo } = datosRubrica.indice_calor;

        let riesgo = 'I'; // Valor por defecto (Riesgo bajo) para temperaturas fuera del rango mínimo

        // Lógica de búsqueda en la matriz
        try {
            // Encontrar el índice de columna para la Temperatura
            // Busca la columna cuyo valor sea el más cercano, pero no mayor, a la temp ingresada.
            let colIndex = temperaturas.findIndex(t => t >= temp);
            if (colIndex === -1) {
                // Si la temperatura es mayor que el máximo de la tabla (42°C)
                colIndex = temperaturas.length - 1; 
            } else if (colIndex > 0 && temperaturas[colIndex] > temp) {
                // Si la temperatura está entre dos valores de la tabla, usa la columna anterior
                colIndex--;
            }

            // Encontrar el índice de fila para la Humedad
            let rowIndex = humedad_relativa.findIndex(h => h >= hum);
            if (rowIndex === -1) {
                // Si la humedad es mayor que el máximo (100%)
                rowIndex = humedad_relativa.length - 1; 
            } else if (rowIndex > 0 && humedad_relativa[rowIndex] > hum) {
                // Si la humedad está entre dos valores de la tabla, usa la fila anterior
                rowIndex--;
            }

            // Asegurar que los índices estén dentro del rango de la matriz (cubre casos < 27°C o < 40%)
            if (colIndex >= 0 && rowIndex >= 0 && temperaturas[colIndex] >= 27 && humedad_relativa[rowIndex] >= 40) {
                 riesgo = matriz_riesgo[rowIndex][colIndex];
            } else {
                // Si está por debajo de los valores mínimos de la tabla (27°C, 40%)
                riesgo = 'I'; 
            }

        } catch (e) {
            console.error("Error al acceder a la matriz:", e);
            riesgo = 'N/A';
        }

        mostrarResultados(riesgo, temp, hum);
    }

    // 3. Mostrar resultados y recomendaciones
    function mostrarResultados(riesgo, temp, hum) {
        const infoRiesgo = datosRubrica.riesgos[riesgo];
        const nombre = document.getElementById('nombre').value || 'N/A';
        const bravo = document.getElementById('bravo').value || 'N/A';

        // Actualizar el estilo y texto del nivel de riesgo
        nivelRiesgoSpan.textContent = riesgo;
        nivelRiesgoSpan.className = `risk-level level-${riesgo}`;
        
        // Actualizar efectos y medidas
        efectosRiesgoP.textContent = infoRiesgo.efectos;
        medidasRiesgoUl.innerHTML = ''; // Limpiar lista
        infoRiesgo.medidas.forEach(medida => {
            const li = document.createElement('li');
            li.innerHTML = medida; // Usamos innerHTML para permitir el tag <strong> si existe
            medidasRiesgoUl.appendChild(li);
        });

        // Mostrar la sección de resultados
        resultadosDiv.classList.remove('hidden');

        // Configurar los manejadores de descarga con todos los datos
        configurarDescarga({
            nombre,
            bravo,
            temperatura: temp,
            humedad: hum,
            nivel: riesgo,
            efectos: infoRiesgo.efectos,
            medidas: infoRiesgo.medidas
        });
    }

    
    // REEMPLAZAR la antigua función configurarDescarga por esta:
    function configurarDescarga(datos) {
    
        // Configuramos el botón único para descargar el DOCX
        descargarWORDBtn.onclick = () => {
            const contenidoHTML = generarContenidoReporte(datos);
         
            // 1. Crear el Blob con tipo MIME para forzar la apertura en Word/procesador de texto
            const blob = new Blob(['\ufeff', contenidoHTML], { 
                type: 'application/msword;charset=utf-8' 
            });
         
            // 2. Definir el nombre del archivo
            const filename = `Reporte_Calor_Riesgo_Nivel_${datos.nivel}.doc`; // Usamos .doc para mejor compatibilidad

            // 3. Crear un enlace de descarga temporal y simular el clic
            const element = document.createElement('a');
            element.setAttribute('href', URL.createObjectURL(blob));
            element.setAttribute('download', filename);
            element.style.display = 'none';
            document.body.appendChild(element);
            element.click();
            document.body.removeChild(element);
        
            alert(`Se ha descargado el reporte como ${filename}. Al abrir, el procesador de texto (Word) interpretará el contenido.`);
        };
    
    }

    // 5. Generar Contenido del Reporte
    function generarContenidoReporte(datos) {
        // Usamos Template Literals para estructurar el HTML de manera legible.
        const htmlContent = `
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <title>Reporte de Índice de Calor</title>
            <style>
                body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 40px; }
                h1 { color: #e67e22; border-bottom: 2px solid #ccc; padding-bottom: 5px; }
                h2 { color: #34495e; margin-top: 25px; }
                .data-section { border: 1px solid #f0f0f0; padding: 15px; margin-bottom: 20px; background-color: #f9f9f9; }
                .risk-level { 
                    font-size: 1.5em; 
                    font-weight: bold; 
                    /* Color dinámico para mejor visualización en el reporte */
                    color: ${datos.nivel === 'IV' ? '#e74c3c' : (datos.nivel === 'III' ? '#e67e22' : (datos.nivel === 'II' ? '#f1c40f' : '#2ecc71'))};
                    padding: 5px;
                }
                ul { list-style-type: disc; margin-left: 20px; }
                li { margin-bottom: 5px; }
                strong { font-weight: bold; }
            </style>
        </head>
        <body>
            <h1>Reporte de Índice de Calor y Riesgo</h1>
            <p><strong>Fecha y Hora:</strong> ${new Date().toLocaleString()}</p>
        
            <h2>Datos de Entrada</h2>
            <div class="data-section">
                <p><strong>Nombre del Empleado:</strong> ${datos.nombre}</p>
                <p><strong>Ubicación/BRAVO:</strong> ${datos.bravo}</p>
                <p><strong>Temperatura Registrada:</strong> ${datos.temperatura}°C</p>
                <p><strong>Humedad Relativa Registrada:</strong> ${datos.humedad}%</p>
            </div>

            <h2>Resultado del Cálculo</h2>
            <div class="data-section">
                <p><strong>NIVEL DE RIESGO:</strong> <span class="risk-level">Nivel ${datos.nivel}</span></p>
                <h3>Efectos en la Salud:</h3>
                <p>${datos.efectos}</p>
            </div>

            <h2>Medidas de Prevención y Protección</h2>
            <ul>
                ${datos.medidas.map(m => `<li>${m}</li>`).join('')}
            </ul>
        
            <p style="margin-top: 30px; font-size: 0.8em; color: #999;">Generado automáticamente por el Sistema de Cálculo de Índice de Calor (IMN).</p>
        </body>
        </html>
        `;
        return htmlContent;
    }

    // Inicialización: Cargar datos y asignar evento
    cargarRubrica();
    calcularBtn.addEventListener('click', calcularRiesgo);
});