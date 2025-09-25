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
        ["DESTINO"]       = new[] {"DESTINO","DEST.","DESTINATION","DESTINO FINAL"},
        ["SALIDA TOPE"]   = new[] {"SALIDA TOPE","TOPE SALIDA","SALIDA_TOPE","TOPE"}
    };

    public static List<LoadRow> ImportExcel(string path)
    {
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.First();

        var lastRow = ws.LastRowUsed().RowNumber();
        var lastCol = ws.LastColumnUsed().ColumnNumber();

        int headerRow = 1, bestHits = -1;
        Dictionary<string,int> colMap = new();

        for (int r = 1; r <= Math.Min(100, lastRow); r++)
        {
            var temp = new Dictionary<string,int>();
            int hits = 0;
            for (int c = 1; c <= lastCol; c++)
            {
                var raw = ws.Cell(r, c).GetString();
                if (string.IsNullOrWhiteSpace(raw)) continue;
                var val = Normalize(raw);
                foreach (var kv in Syns)
                {
                    if (temp.ContainsKey(kv.Key)) continue;
                    if (kv.Value.Select(Normalize).Contains(val))
                    {
                        temp[kv.Key] = c; hits++;
                    }
                }
            }
            if (hits > bestHits) { bestHits = hits; headerRow = r; colMap = temp; }
        }

        string GetText(int row, string key)
        {
            if (!colMap.TryGetValue(key, out var c)) return "";
            return ws.Cell(row, c).GetString();
        }

        TimeSpan? GetTime(int row, string key)
        {
            if (!colMap.TryGetValue(key, out var c)) return null;
            var cell = ws.Cell(row, c);

            if (cell.DataType == XLDataType.Number)
            {
                var d = cell.GetDouble(); var ts = TimeSpan.FromDays(d);
                return new TimeSpan(ts.Hours, ts.Minutes, 0);
            }
            if (cell.DataType == XLDataType.DateTime)
            {
                var dt = cell.GetDateTime(); return new TimeSpan(dt.Hour, dt.Minute, 0);
            }
            var s = cell.GetString();
            if (string.IsNullOrWhiteSpace(s)) return null;
            if (TimeSpan.TryParse(s, out var t1)) return new TimeSpan(t1.Hours, t1.Minutes, 0);

            var cleaned = s.Replace(".", ":").Replace(",", ":");
            if (cleaned.Length is 3 or 4) { cleaned = cleaned.PadLeft(4,'0'); cleaned = cleaned.Insert(2, ":"); }
            if (TimeSpan.TryParse(cleaned, out var t2)) return new TimeSpan(t2.Hours, t2.Minutes, 0);

            return null;
        }

        var list = new List<LoadRow>();
        for (int r = headerRow + 1; r <= lastRow; r++)
        {
            bool empty = true;
            for (int c = 1; c <= lastCol; c++)
                if (!string.IsNullOrWhiteSpace(ws.Cell(r,c).GetString())) { empty = false; break; }
            if (empty) continue;

            list.Add(new LoadRow
            {
                Transportista = GetText(r, "TRANSPORTISTA"),
                Matricula     = GetText(r, "MATRICULA"),
                Destino       = GetText(r, "DESTINO"),
                Muelle        = GetText(r, "MUELLE"),
                Estado        = GetText(r, "ESTADO"),
                SalidaTope    = GetTime(r, "SALIDA TOPE")
            });
        }

        return list.OrderBy(x => x.Muelle).ToList();
    }

    private static string Normalize(string s) =>
        (s ?? string.Empty).Trim().ToUpperInvariant()
          .Replace("Á","A").Replace("É","E").Replace("Í","I")
          .Replace("Ó","O").Replace("Ú","U").Replace("Ü","U").Replace("Ñ","N");
}
