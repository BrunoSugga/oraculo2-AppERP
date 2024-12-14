using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace AppERP.Services
{
    public class ExcelService
    {
        private List<Dictionary<string, object>> data = new List<Dictionary<string, object>>();
        private List<string> sheetNames = new List<string>();

        public ExcelService(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    throw new Exception("El archivo Excel no se encuentra en la ruta especificada.");
                }

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        throw new Exception("El archivo Excel no contiene ninguna hoja de trabajo.");
                    }

                    foreach (var ws in package.Workbook.Worksheets)
                    {
                        sheetNames.Add(ws.Name);
                    }

                    if (!sheetNames.Contains("Informe"))
                    {
                        throw new Exception("La hoja de trabajo 'Informe' no se encontró en el archivo Excel.");
                    }

                    var worksheet = package.Workbook.Worksheets["Informe"];
                    var rowCount = worksheet.Dimension.Rows;
                    var colCount = worksheet.Dimension.Columns;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var rowData = new Dictionary<string, object>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            rowData[worksheet.Cells[1, col].Text] = worksheet.Cells[row, col].Text;
                        }
                        data.Add(rowData);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al leer el archivo Excel: {ex.Message}"); // Añadir una línea de depuración
                throw new Exception($"Error al leer el archivo Excel: {ex.Message}");
            }
        }

        public (List<Dictionary<string, object>> data, List<string> sheetNames) ReadExcel()
        {
            return (data, sheetNames);
        }

        public string AnswerQuestion(string question)
        {
            Console.WriteLine($"Pregunta recibida: {question}"); // Añadir mensaje de depuración
            
            // Procesar la pregunta e interpretar qué datos buscar en el Excel
            if (question.Contains("ventas", StringComparison.OrdinalIgnoreCase) && 
                (question.Contains("total", StringComparison.OrdinalIgnoreCase) || question.Contains("totales", StringComparison.OrdinalIgnoreCase)))
            {
                var sales = data.Where(item => item["Documento"].ToString().Contains("credito", StringComparison.OrdinalIgnoreCase) || 
                                               item["Documento"].ToString().Contains("contado", StringComparison.OrdinalIgnoreCase))
                                .Sum(item => Convert.ToDouble(item["TotalRenglonPesos"]));
                return $"Las ventas totales son: {sales}";
            }
            else if (question.Contains("ventas", StringComparison.OrdinalIgnoreCase) && 
                     question.Contains("artículo", StringComparison.OrdinalIgnoreCase))
            {
                // Lógica para responder preguntas sobre ventas por artículo
                var ventasPorArticulo = data.GroupBy(item => item["Artículo"].ToString())
                                            .Select(g => new { Artículo = g.Key, TotalVentas = g.Sum(item => Convert.ToDouble(item["TotalRenglonPesos"])) });

                return $"Ventas por artículo:\n{string.Join("\n", ventasPorArticulo.Select(v => $"{v.Artículo}: {v.TotalVentas}"))}";
            }
            else if (question.Contains("clientes", StringComparison.OrdinalIgnoreCase))
            {
                // Lógica para responder preguntas sobre ventas por clientes
                var ventasPorCliente = data.GroupBy(item => item["Cliente"].ToString())
                                           .Select(g => new { Cliente = g.Key, TotalVentas = g.Sum(item => Convert.ToDouble(item["TotalRenglonPesos"])) });

                return $"Ventas por cliente:\n{string.Join("\n", ventasPorCliente.Select(v => $"{v.Cliente}: {v.TotalVentas}"))}";
            }

            return "No se pudo interpretar la pregunta.";
        }
    }
}
