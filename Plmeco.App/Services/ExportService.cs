using System.Collections.Generic;
using ClosedXML.Excel;
using Plmeco.App.Models;

namespace Plmeco.App.Services
{
    public static class ExportService
    {
        // Exporta la pestaña actual a un .xlsx sencillo
        public static void ExportExcel(string path, IList<LoadRow> rows)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Cargas");

            // Cabeceras
            int r = 1;
            ws.Cell(r,1).Value = "ESTADO";
            ws.Cell(r,2).Value = "PRECINTO";
            ws.Cell(r,3).Value = "LLEGADA REAL";
            ws.Cell(r,4).Value = "SALIDA REAL";
            ws.Cell(r,5).Value = "LEX";
            ws.Cell(r,6).Value = "INCIDENCIAS";
            ws.Cell(r,7).Value = "DESTINO";
            ws.Cell(r,8).Value = "SALIDA TOPE";

            foreach (var c in Enumerable.Range(1,8)) ws.Cell(1,c).Style.Font.Bold = true;

            // Filas
            foreach (var x in rows)
            {
                r++;
                ws.Cell(r,1).Value = x.Estado;
                ws.Cell(r,2).Value = x.Precinto;
                ws.Cell(r,3).Value = Format(x.LlegadaReal);
                ws.Cell(r,4).Value = Format(x.SalidaReal);
                ws.Cell(r,5).Value = x.Lex;
                ws.Cell(r,6).Value = x.Incidencias;
                ws.Cell(r,7).Value = x.Destino;
                ws.Cell(r,8).Value = Format(x.SalidaTope);
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(path);

            static string Format(System.TimeSpan? ts) =>
                ts.HasValue ? ts.Value.Hours.ToString("00")+":"+ts.Value.Minutes.ToString("00") : "";
        }
    }
}
