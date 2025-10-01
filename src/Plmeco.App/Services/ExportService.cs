using System.Collections.Generic;
using ClosedXML.Excel;
using Plmeco.App.Models;

namespace Plmeco.App.Services
{
    public static class ExportService
    {
        public static void ExportExcel(string filePath, IEnumerable<LoadRow> rows)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Cargas");

            // Cabeceras (incluye PRECINTO entre ESTADO y LLEGADA REAL)
            string[] headers = {
                "MATRICULA", "DESTINO", "MUELLE", "ESTADO",
                "PRECINTO",
                "LLEGADA REAL", "SALIDA REAL", "SALIDA TOPE",
                "INCIDENCIAS", "LEX"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontColor = XLColor.Black;
            }

            int r = 2;
            foreach (var x in rows)
            {
                // Columnas
                ws.Cell(r, 1).Value = x.Matricula;                                            // MATRICULA
                ws.Cell(r, 2).Value = x.Destino;                                              // DESTINO
                ws.Cell(r, 3).Value = x.Muelle;                                               // MUELLE
                ws.Cell(r, 4).Value = x.Estado;                                               // ESTADO
                ws.Cell(r, 5).Value = x.Precinto;                                             // PRECINTO (NUEVO)
                ws.Cell(r, 6).Value = x.LlegadaReal?.ToString(@"hh\:mm") ?? "";               // LLEGADA REAL
                ws.Cell(r, 7).Value = x.SalidaReal?.ToString(@"hh\:mm") ?? "";                // SALIDA REAL
                ws.Cell(r, 8).Value = x.SalidaTope?.ToString(@"hh\:mm") ?? "";                // SALIDA TOPE
                ws.Cell(r, 9).Value = x.Incidencias;                                          // INCIDENCIAS
                ws.Cell(r,10).Value = x.Lex ? "✔" : "";                                       // LEX

                // Fuente negra + negrita por defecto en toda la fila
                for (int c = 1; c <= 10; c++)
                {
                    ws.Cell(r, c).Style.Font.FontColor = XLColor.Black;
                    ws.Cell(r, c).Style.Font.Bold = true;
                }

                // ===== Colores (mismos criterios que la app) =====
                // DESTINO verde si ESTADO = OK
                if (!string.IsNullOrWhiteSpace(x.Estado) &&
                    x.Estado.Equals("OK", System.StringComparison.OrdinalIgnoreCase))
                {
                    ws.Cell(r, 2).Style.Fill.BackgroundColor = XLColor.LightGreen;
                }

                // ESTADO naranja si CARGANDO
                if (!string.IsNullOrWhiteSpace(x.Estado) &&
                    x.Estado.Equals("CARGANDO", System.StringComparison.OrdinalIgnoreCase))
                {
                    ws.Cell(r, 4).Style.Fill.BackgroundColor = XLColor.Orange;
                }

                // Fila roja si ANULADO
                if (!string.IsNullOrWhiteSpace(x.Estado) &&
                    x.Estado.Equals("ANULADO", System.StringComparison.OrdinalIgnoreCase))
                {
                    ws.Row(r).Style.Fill.BackgroundColor = XLColor.LightPink;
                    ws.Row(r).Style.Font.FontColor = XLColor.Black; // mantenemos negro para legibilidad
                }

                // Fila verde si LEX (si coincide con ANULADO, prevalece el de arriba)
                if (x.Lex)
                {
                    ws.Row(r).Style.Fill.BackgroundColor = XLColor.LightGreen;
                    ws.Row(r).Style.Font.FontColor = XLColor.Black;
                }

                // INCIDENCIAS naranja si NO vacío
                if (!string.IsNullOrWhiteSpace(x.Incidencias))
                {
                    ws.Cell(r, 9).Style.Fill.BackgroundColor = XLColor.Orange;
                }

                r++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(filePath);
        }
    }
}
