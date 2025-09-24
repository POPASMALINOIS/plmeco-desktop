using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using Plmeco.App.Models;

namespace Plmeco.App.Services;

public static class ImportService
{
    private static readonly Dictionary<string,string[]> Syns = new()
    {
        ["TRANSPORTISTA"] = new[] {"TRANSPORTISTA","TRANSPORTE","CARRIER","TRANSPORTER"},
        ["MATRICULA"]     = new[] {"MATRICULA","MATRÍCULA","PLACA","MATRIC","REG","REGISTRO"},
        ["MUELLE"]        = new[] {"MUELLE","MUEL.","DOCK"},
        ["ESTADO"]        = new[] {"ESTADO","STATUS","EST."},
        ["DESTINO"]       = new[] {"DESTINO","DEST.","DESTINATION"},
        ["SALIDA TOPE"]   = new[] {"SALIDA TOPE","TOPE SALIDA","SALIDA_TOPE","SALIDA TOPE"} // ojo espacios raros
    };

    public static List<LoadRow> ImportExcel(string path)
    {
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.First();

        var lastRow = ws.LastRowUsed().RowNumber();
        var lastCol = ws.LastColumnUsed().ColumnNumber();

        // 1) Detectar fila de encabezados: buscamos la que más sinónimos contiene en las primeras 100 filas
        int headerRow = 1, bestHits = -1;
        Dictionary<string,int> colMap = new();

        for (int r = 1; r <= Math.Min(100, lastRow); r++)
        {
            var temp = new Dictionary<string,int>();
            int hits = 0;

            for (int c = 1; c <= lastCol; c++)
            {
                var val = (ws.Cell(r, c).GetString() ?? "").Trim();
                if (string.IsNullOrEmpty(val)) continue;
                var norm = Normalize(val);

                foreach (var kv in Syns)
                {
                    if (temp.ContainsKey(kv.Key)) continue;
                    if (kv.Value.Any(s => Normalize(s) == norm))
                    {
                        temp[kv.Key] = c;
                        hits++;
                    }
                }
            }

            if (hits > bestHits)
            {
                bestHits = hits;
                headerRow = r;
                colMap = temp;
            }
        }

        // 2) Funciones auxiliares
        string GetCell(int row, string key)
            => colMap.TryGetValue(key, out var c) ? ws.Cell(row, c).GetString() : "";

        TimeSpan? ParseTime(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            if (TimeSpan.TryParse(s, out var t)) return t;
            if (double.TryParse(s, out var d)) return TimeSpan.FromDays(d); // si viene como número Excel
            return null;
        }

        // 3) Leer filas de datos
        var list = new List<LoadRow>();
        for (int r = headerRow + 1; r <= lastRow; r++)
        {
            // Si la fila está completamente vacía, saltamos
            bool vacia = true;
            for (int c = 1; c <= lastCol; c++)
                if (!string.IsNullOrWhiteSpace(ws.Cell(r, c).GetString())) { vacia = false; break; }
            if (vacia) continue;

            list.Add(new LoadRow
            {
                Transportista = GetCell(r, "TRANSPORTISTA"),
                Matricula     = GetCell(r, "MATRICULA"),
                Destino       = GetCell(r, "DESTINO"),
                Muelle        = GetCell(r, "MUELLE"),
                Estado        = GetCell(r, "ESTADO"),
                SalidaTope    = ParseTime(GetCell(r, "SALIDA TOPE"))
            });
        }

        // 4) Orden por muelle
        return list.OrderBy(x => x.Muelle).ToList();
    }

    private static string Normalize(string s) =>
        s.ToUpperInvariant()
         .Replace("Á","A").Replace("É","E").Replace("Í","I")
         .Replace("Ó","O").Replace("Ú","U").Replace("Ü","U")
         .Replace("Ñ","N").Trim();
}
