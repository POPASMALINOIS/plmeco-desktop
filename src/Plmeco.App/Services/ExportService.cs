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
            string[] headers = {
                "MATRICULA","DESTINO","MUELLE","ESTADO",
                "LLEGADA REAL","SALIDA REAL","SALIDA TOPE",
                "INCIDENCIAS","LEX"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
            }

            int r = 2;
            foreach (var x in rows)
            {
                ws.Cell(r, 1).Value = x.Matricula;
                ws.Cell(r, 2).Value = x.Destino;
                ws.Cell(r, 3).Value = x.Muelle;
                ws.Cell(r, 4).Value = x.Estado;

                // Horas: guardamos como texto hh:mm (simple y robusto)
                ws.Cell(r, 5).Value = x.LlegadaReal.HasValue ? x.LlegadaReal.Value.ToString(@"hh\:mm") : "";
                ws.Cell(r, 6).Value = x.SalidaReal.HasValue  ? x.SalidaReal.Value.ToString(@"hh\:mm")  : "";
                ws.Cell(r, 7).Value = x.SalidaTope.HasValue  ? x.SalidaTope.Value.ToString(@"hh\:mm")  : "";

                ws.Cell(r, 8).Value = x.Incidencias;
                ws.Cell(r, 9).Value = x.Lex ? "✔" : "";

                // ==== Colores (mismos criterios que la app) ====

                // DESTINO verde si ESTADO = OK (case-insensitive)
                if (!string.IsNullOrWhiteSpace(x.Estado) &&
                    x.Estado.Equals("OK", System.StringComparison.OrdinalIgnoreCase))
                {
                    ws.Cell(r, 2).Style.Fill.BackgroundColor = XLColor.LightGreen;
                    ws.Cell(r, 2).Style.Font.Bold = true;
                }

                // ESTADO naranja si CARGANDO
                if (!string.IsNullOrWhiteSpace(x.Estado) &&
                    x.Estado.Equals("CARGANDO", System.StringComparison.OrdinalIgnoreCase))
                {
                    ws.Cell(r, 4).Style.Fill.BackgroundColor = XLColor.LightOrange; // <— CORRECTO
                    ws.Cell(r, 4).Style.Font.Bold = true;
                }

                // Fila roja si CAMION ANULADO
                if (!string.IsNullOrWhiteSpace(x.Estado) &&
                    x.Estado.Equals("CAMION ANULADO", System.StringComparison.OrdinalIgnoreCase))
                {
                    ws.Row(r).Style.Fill.BackgroundColor = XLColor.LightPink;
                    ws.Row(r).Style.Font.FontColor = XLColor.DarkRed;
                    ws.Row(r).Style.Font.Bold = true;
                }

                // Fila verde si LEX (si coincide con ANULADO, arriba ya lo pintamos en rojo)
                if (x.Lex)
                {
                    ws.Row(r).Style.Fill.BackgroundColor = XLColor.LightGreen;
                }

                // INCIDENCIAS naranja si NO vacío
                if (!string.IsNullOrWhiteSpace(x.Incidencias))
                {
                    ws.Cell(r, 8).Style.Fill.BackgroundColor = XLColor.LightOrange; // <— CORRECTO
                    ws.Cell(r, 8).Style.Font.Bold = true;
                }

                r++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(filePath);
        }
    }
}
