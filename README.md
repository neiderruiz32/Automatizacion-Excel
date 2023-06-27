# Automatización de actualización de datos en Excel desde un archivo de texto

Este programa en C# automatiza la actualización de datos en un archivo de Excel a partir de un archivo de texto. El programa monitorea continuamente el archivo de texto en busca de modificaciones y agrega los datos nuevos o modificados al archivo de Excel.

## Requisitos

- Microsoft Office Excel instalado en el sistema.
- Archivo de texto con los datos a importar.
- Ruta del archivo de texto válido.

## Funcionamiento

El programa sigue los siguientes pasos para realizar la automatización:

1. Se establece la ruta del archivo de texto que contiene los datos a importar.
2. Se crea una instancia de la aplicación Excel y se obtiene una referencia al libro de trabajo abierto o se crea uno nuevo.
3. Se encuentra la última fila con datos en la columna A del archivo de Excel.
4. Se inicia un bucle infinito para monitorear el archivo de texto en busca de modificaciones.
5. Si el archivo de texto ha sido modificado después de la última vez que se procesó, se procede a realizar la actualización.
6. Se abre el archivo de texto en modo lectura y se posiciona el cursor en la última posición conocida.
7. Se leen las líneas nuevas o modificadas desde la última línea procesada.
8. Se verifica cada línea para asegurarse de que contenga los datos esperados.
9. Si la línea es válida, se divide por las comas y se agregan los datos en la siguiente fila de la hoja de trabajo del archivo de Excel.
10. Se actualiza la última posición conocida en el archivo de texto y se guarda la fecha y hora de la última modificación.
11. Se guarda el archivo de Excel.
12. El programa se pausa durante un segundo y luego continúa monitoreando el archivo de texto.

## Instrucciones de uso

1. Abre el archivo de texto con los datos que deseas importar en la ruta especificada por `rutaArchivoTexto`.
2. Ejecuta el programa.
3. Mantén abiertos tanto el archivo de texto como el archivo de Excel donde se realizarán las actualizaciones.
4. Guarda el archivo de texto cada vez que realices modificaciones para que los cambios se reflejen en el archivo de Excel.

**Nota:** Si no se encuentra una instancia de Excel abierta al ejecutar el programa, se creará una nueva instancia automáticamente.

## Importante

Es importante asegurarse de que los datos en el archivo de texto estén correctamente formateados y cumplan con las expectativas del programa. Cada línea del archivo de texto debe contener 12 elementos separados por comas, que serán agregados en las columnas correspondientes del archivo de Excel.
En caso de que ocurra algún error durante la automatización, se mostrará un mensaje de error en la consola.
Recuerda liberar los recursos de Excel al finalizar la ejecución del programa.
