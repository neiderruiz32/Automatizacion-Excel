using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static void Main()
    {
        // Ruta del archivo de texto
        string rutaArchivoTexto = @"C:\Users\PC\OneDrive\Documentos\Escritorio\Actividad\Actividad.txt";

        // Crea una instancia de Excel.Application
        Excel.Application excelApp = null;
        Excel.Workbook workbook = null;
        Excel.Worksheet worksheet = null;

        try
        {
            try
            {
                // Intenta obtener la instancia de Excel.Application
                excelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Si no se encuentra una instancia de Excel abierta, crea una nueva instancia
                excelApp = new Excel.Application();
            }

            // Establece la visibilidad de la aplicación Excel
            excelApp.Visible = true;

            // Obtiene una referencia al libro de trabajo abierto o crea uno nuevo
            if (excelApp.Workbooks.Count > 0)
            {
                workbook = excelApp.ActiveWorkbook;
            }
            else
            {
                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.ActiveSheet; // Agregar esta línea para obtener la referencia a la hoja de trabajo activa
            }

            // Si no se ha asignado una hoja de trabajo, obtener la referencia a la primera hoja de trabajo
            if (worksheet == null)
            {
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            }

            // Encuentra la última fila con datos en la columna A
            int ultimaFila = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;

            Console.WriteLine("La automatización está en ejecución. Mantén ambos archivos abiertos y guarda el archivo de texto para que los cambios se actualicen en el archivo Excel.");

            DateTime ultimaModificacion = DateTime.Now;
            long ultimaPosicion = 0;

            while (true)
            {
                // Verifica si el archivo de texto ha sido modificado después de la última vez que se procesó
                if (File.GetLastWriteTime(rutaArchivoTexto) > ultimaModificacion)
                {
                    // Abre el archivo de texto en modo lectura
                    using (var stream = new FileStream(rutaArchivoTexto, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        // Posiciona el cursor en la última posición conocida
                        stream.Position = ultimaPosicion;

                        // Lee solo las líneas nuevas o modificadas desde la última línea procesada
                        using (var reader = new StreamReader(stream))
                        {
                            string linea;
                            while ((linea = reader.ReadLine()) != null)
                            {
                                // Verifica si la línea no está vacía
                                if (!string.IsNullOrEmpty(linea))
                                {
                                    // Divide la línea por las comas
                                    string[] datos = linea.Split(',');

                                    // Verifica si la línea contiene los datos esperados
                                    if (datos.Length == 12)
                                    {
                                        // Agrega cada dato en la siguiente fila de la hoja de trabajo
                                        for (int i = 0; i < datos.Length; i++)
                                        {
                                            // Establece el valor en la celda correspondiente
                                            worksheet.Cells[ultimaFila, i + 1].Value = datos[i];
                                        }

                                        // Incrementa la última fila por la cantidad de datos agregados
                                        ultimaFila++;
                                    }
                                }
                            }

                            // Actualiza la última posición conocida
                            ultimaPosicion = stream.Position;
                        }
                    }

                    // Actualiza la fecha y hora de la última modificación
                    ultimaModificacion = File.GetLastWriteTime(rutaArchivoTexto);

                    // Guarda el archivo Excel
                    workbook.Save();
                }

                // Pausa el programa durante un segundo
                System.Threading.Thread.Sleep(1000);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Ocurrió un error durante la automatización: " + ex.Message);
        }
        finally
        {
            // Liberar recursos de Excel
            if (worksheet != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            if (workbook != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            if (excelApp != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }

            worksheet = null;
            workbook = null;
            excelApp = null;

            GC.Collect();
        }
    }
}
