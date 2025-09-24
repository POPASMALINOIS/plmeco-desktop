using ClosedXML.Excel;
using Plmeco.App.Models;

namespace Plmeco.App.Services;

public static class ImportService
{
    private static readonly Dictionary<string,string[]> Syns = new()
    {
        ["TRANSPORTISTA"] = new[] {"TRANSPORTISTA","TRANSPORTE","CARRIER"},
        ["MATRICULA"]     = new[] {"MATRICULA","MATRÍCULA","PLACA","MATRIC","REG"},
        ["MUELLE"]        = new[] {"MUELLE","MUEL.","DOCK"},
        ["ESTADO"]        = new[] {"ESTADO","STATUS","EST."},
        ["DESTINO"]       = new[] {"DESTINO","DEST.","DESTINATION"},
        ["SALIDA TOPE"]   = new[] {"SALIDA TOPE","TOPE SALIDA","SALIDA_TOPE"}
    };

    public static List<LoadRow> ImportExcel(string path)
    {
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.First();

        var lastRow = ws.LastRowUsed().RowNumber();
        var lastCol = ws.LastColumnUsed().ColumnNumber();

        // fila de encabezados = 1 por simplicidad
        var headerRow = 1;
        var list = new List<LoadRow>();

        for (int r = headerRow + 1; r <= lastRow; r++)
        {
            string Get(string key) => ws.Cell(r, Syns[key].SelectMany(s => Enumerable.Range(1,lastCol)
                .Where(c => Normalize(ws.Cell(headerRow,c).GetString()) == Normalize(s)))
                .FirstOrDefault()).GetString();

            TimeSpan? ParseTime(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return null;
                if (TimeSpan.TryParse(s, out var t)) return t;
                if (double.TryParse(s, out var d)) return TimeSpan.FromDays(d);
                return null;
            }

            list.Add(new LoadRow {
                Transportista = Get("TRANSPORTISTA"),
                Matricula     = Get("MATRICULA"),
                Destino       = Get("DESTINO"),
                Muelle        = Get("MUELLE"),
                Estado        = Get("ESTADO"),
                SalidaTope    = ParseTime(Get("SALIDA TOPE"))
            });
        }

        return list.OrderBy(x => x.Muelle).ToList();
    }

    private static string Normalize(string s) =>
        s.ToUpperInvariant().Replace("Á","A").Replace("É","E").Replace("Í","I")
                             .Replace("Ó","O").Replace("Ú","U").Replace("Ü","U")
                             .Replace("Ñ","N").Trim();
}
