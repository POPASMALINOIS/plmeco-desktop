using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;
using Plmeco.App.Models;

namespace Plmeco.App.Services
{
    public static class ImportService
    {
        public sealed class ImportResult
        {
            public string SheetName { get; init; } = "Hoja";
            public int HeaderRow { get; init; }
            public List<LoadRow> Rows { get; init; } = new();
            public Dictionary<string,int> Map { get; init; } = new();
        }

        private static readonly Dictionary<string, string[]> Syns = new()
        {
            ["TRANSPORTISTA"] = new[] { "TRANSPORTISTA","TRANSPORTE","CARRIER","TRANSPORTER" },
            ["MATRICULA"]     = new[] { "MATRICULA","MATRÍCULA","PLACA","REG","REGISTRO" },
            ["MUELLE"]        = new[] { "MUELLE","MUEL.","DOCK" },
            ["ESTADO"]        = new[] { "ESTADO","STATUS","EST." },
            ["DESTINO"]       = new[] { "DESTINO","DEST.","DESTINATION","DESTINO FINAL" },
            ["SALIDA TOPE"]   = new[] { "SALIDA TOPE","TOPE SALIDA","SALIDA_TOPE","TOPE","TOPE DE SALIDA" },
            ["PRECINTO"]      = new[] { "PRECINTO","SEAL","SEALS","Nº PRECINTO","NUM PRECINTO","NUMERO PRECINTO" }
        };

        private static string N(string s) =>
            (s ?? string.Empty).Trim().ToUpperInvariant()
                .Replace("Á","A").Replace("É","E").Replace("Í","I")
                .Replace("Ó","O").Replace("Ú","U").Replace("Ü","U").Replace("Ñ","N");

        public static ImportResult ImportExcel(string path)
        {
            using var wb = new XLWorkbook(path);

            IXLWorksheet? bestWs = null;
            int bestHits = -1;
            int bestHeaderRow = 1;
            Dictionary<string,int> bestMap = new();

            foreach (var ws in wb.Worksheets)
            {
                var (hits, headerRow, map) = DetectHeaders(ws);
                if (hits > bestHits)
                {
                    bestHits = hits;
                    bestHeaderRow = headerRow;
                    bestMap = map;
                    bestWs = ws;
                }
            }

            var result = new ImportResult();
            if (bestWs == null || !HasMinimum(bestMap))
                return result; // vacío (UI avisará)

            result.SheetName = bestWs.Name;
            result.HeaderRow = bestHeaderRow;
            result.Map = bestMap;

            var lastRow = bestWs.LastRowUsed()?.RowNumber() ?? 0;
            var lastCol = bestWs.LastColumnUsed()?.ColumnNumber() ?? 0;

            string GetText(int r, string key)
            {
                if (!bestMap.TryGetValue(key, out var c)) return "";
                return bestWs.Cell(r, c).GetString();
            }

            TimeSpan? GetTime(int r, string key)
            {
                if (!bestMap.TryGetValue(key, out var c)) return null;
                var cell = bestWs.Cell(r, c);

                if (cell.DataType == XLDataType.Number)
                {
                    var d = cell.GetDouble(); // días
                    var ts = TimeSpan.FromDays(d);
                    return new TimeSpan(ts.Hours, ts.Minutes, 0);
                }
                if (cell.DataType == XLDataType.DateTime)
                {
                    var dt = cell.GetDateTime();
                    return new TimeSpan(dt.Hour, dt.Minute, 0);
                }

                var s = cell.GetString()?.Trim();
                if (string.IsNullOrEmpty(s)) return null;

                s = s.Replace(".", ":").Replace(",", ":");
                if (TimeSpan.TryParse(s, CultureInfo.InvariantCulture, out var t1))
                    return new TimeSpan(t1.Hours, t1.Minutes, 0);

                var digits = new string(s.Where(char.IsDigit).ToArray());
                if (digits.Length is 3 or 4)
                {
                    digits = digits.PadLeft(4, '0');
                    var fixedStr = digits.Insert(2, ":");
                    if (TimeSpan.TryParse(fixedStr, CultureInfo.InvariantCulture, out var t2))
                        return new TimeSpan(t2.Hours, t2.Minutes, 0);
                }
                return null;
            }

            for (int r = bestHeaderRow + 1; r <= lastRow; r++)
            {
                bool empty = true;
                for (int c = 1; c <= lastCol; c++)
                {
                    if (!string.IsNullOrWhiteSpace(bestWs.Cell(r, c).GetString()))
                    { empty = false; break; }
                }
                if (empty) continue;

                var row = new LoadRow
                {
                    Transportista = GetText(r, "TRANSPORTISTA"),
                    Matricula     = GetText(r, "MATRICULA"),
                    Destino       = GetText(r, "DESTINO"),
                    Muelle        = GetText(r, "MUELLE"),
                    Estado        = GetText(r, "ESTADO"),
                    Precinto      = GetText(r, "PRECINTO"),
                    SalidaTope    = GetTime (r, "SALIDA TOPE")
                };
                result.Rows.Add(row);
            }

            result.Rows = result.Rows
                .OrderBy(x =>
                {
                    var d = new string((x.Muelle ?? "").Where(char.IsDigit).ToArray());
                    return int.TryParse(d, out var n) ? n : int.MaxValue;
                })
                .ThenBy(x => x.Muelle)
                .ToList();

            return result;
        }

        private static (int hits, int headerRow, Dictionary<string,int> map) DetectHeaders(IXLWorksheet ws)
        {
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;

            int bestHits = -1;
            int headerRow = 1;
            var bestMap = new Dictionary<string,int>();

            int maxRow = Math.Min(100, Math.Max(1, lastRow));
            for (int r = 1; r <= maxRow; r++)
            {
                int hits = 0;
                var map = new Dictionary<string,int>();
                for (int c = 1; c <= lastCol; c++)
                {
                    var val = N(ws.Cell(r, c).GetString());
                    if (string.IsNullOrWhiteSpace(val)) continue;

                    foreach (var kv in Syns)
                    {
                        if (map.ContainsKey(kv.Key)) continue;
                        if (kv.Value.Select(N).Contains(val))
                        {
                            map[kv.Key] = c;
                            hits++;
                        }
                    }
                }

                if (hits > bestHits)
                {
                    bestHits = hits;
                    headerRow = r;
                    bestMap = map;
                }
            }

            return (bestHits, headerRow, bestMap);
        }

        private static bool HasMinimum(Dictionary<string,int> map)
        {
            string[] needed = { "MATRICULA","DESTINO","MUELLE","ESTADO" };
            return needed.All(map.ContainsKey);
        }
    }
}
