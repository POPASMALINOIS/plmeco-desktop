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

            // Cabeceras
            string[] headers = { "MATRICULA", "DESTINO", "MUELLE", "ESTADO",
                                 "LLEGADA REAL", "SALIDA REAL", "SALIDA TOPE",
                                 "INCIDENCIAS", "LEX" };

            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(1, i + 1).Value = headers[i];
                ws.Cell(1, i + 1).Style.Font.Bold = true;
            }

            int rowIndex = 2;
            foreach (var row in rows)
            {
                ws.Cell(rowIndex, 1).Value = row.Matricula;
                ws.Cell(rowIndex, 2).Value = row.Destino;
                ws.Cell(rowIndex, 3).Value = row.Muelle;
                ws.Cell(rowIndex, 4).Value = row.Estado;
                ws.Cell(rowIndex, 5).Value = row.LlegadaReal?.ToString(@"hh\:mm");
                ws.Cell(rowIndex, 6).Value = row.SalidaReal?.ToString(@"hh\:mm");
                ws.Cell(rowIndex, 7).Value = row.SalidaTope?.ToString(@"hh\:mm");
                ws.Cell(rowIndex, 8).Value = row.Incidencias;
                ws.Cell(rowIndex, 9).Value = row.Lex ? "✔" : "";

                // --- Colores según reglas ---
                if (row.Estado.Equals("OK", System.StringComparison.OrdinalIgnoreCase))
                    ws.Cell(rowIndex, 2).Style.Fill.BackgroundColor = XLColor.LightGreen;

                if (row.Estado.Equals("CARGANDO", System.StringComparison.OrdinalIgnoreCase))
                    ws.Cell(rowIndex, 4).Style.Fill.BackgroundColor = XLColor.OrangeLight;

                if (row.Estado.Equals("CAMION ANULADO", System.StringComparison.OrdinalIgnoreCase))
                {
                    ws.Row(rowIndex).Style.Fill.BackgroundColor = XLColor.LightPink;
                    ws.Row(rowIndex).Style.Font.FontColor = XLColor.DarkRed;
                }

                if (row.Lex)
                    ws.Row(rowIndex).Style.Fill.BackgroundColor = XLColor.LightGreen;

                if (!string.IsNullOrWhiteSpace(row.Incidencias))
                    ws.Cell(rowIndex, 8).Style.Fill.BackgroundColor = XLColor.OrangeLight;

                rowIndex++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(filePath);
        }
    }
}
