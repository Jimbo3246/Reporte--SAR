# Proyecto de Exportación de Datos

Este proyecto contiene varias funciones para exportar datos de una hoja de cálculo de Google Sheets a diferentes formatos, como Excel y CSV.

## Funciones disponibles

### Eliminar datos de la BD

La función `eliminar_data_BD()` borra el contenido de la hoja "BD" en el libro activo de Google Sheets.

### Exportar a Excel

La función `obtener_Excel()` filtra los datos de la hoja "BD" y copia los resultados en la hoja "Resultado". Luego, guarda la hoja "Resultado" en formato XLSX en Google Drive con un nombre basado en la fecha y hora de ejecución.

### Exportar a CSV

La función `obtener_Csv()` también filtra los datos de la hoja "Resultado" y genera un archivo CSV. El mensaje en cada fila se genera dinámicamente utilizando el número de reclamo almacenado en la columna 37.

## Utilidades

### Funciones de formato de fecha

- `formatDate(date)`: Formatea una fecha en el formato "yyyy-MM-dd HH:mm:ss".
- `formatDateForFileName(date)`: Formatea una fecha en el formato "yyyyMMddHHmmss".

### Función de agregar ceros a la izquierda

- `padZero(value)`: Agrega ceros a la izquierda de un valor numérico si es necesario.

## Notas adicionales

- El archivo de Excel se guarda temporalmente como una copia del libro activo, se convierten las fórmulas en texto y se eliminan las hojas que no sean la hoja "Resultado". Luego, se obtiene el archivo XLSX desde la URL de exportación y se crea como un archivo en Google Drive. Finalmente, se elimina el archivo temporal.
- El archivo CSV se guarda en la carpeta con ID "1n2Vy2qc-bY5P1YKtcAVIqqabC5uS_MZF" en Google Drive con un nombre basado en la fecha y hora de ejecución.

---
**Nota**: Este README.md fue generado automáticamente. Asegúrate de mantenerlo actualizado si realizas cambios en el código.
