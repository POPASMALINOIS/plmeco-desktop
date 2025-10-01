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
        // Sinónimos por columna (case/acento-insensitive)
        private static readonly Dictionary<string, string[]> Syns = new()
        {
            ["TRANSPORTISTA"] = new[] { "TRANSPORTISTA", "TRANSPORTE", "CARRIER", "TRANSPORTER" },
            ["MATRICULA"]     = new[] { "MATRICULA", "MATRÍCULA", "PLACA", "MATRIC", "REG", "REGISTRO" },
            ["MUELLE"]        = new[] { "MUELLE", "MUEL.", "DOCK" },
            ["ESTADO"]        = new[] { "ESTADO", "STATUS", "EST." },
            ["DESTINO"]       = new[] { "DESTINO", "DEST.", "DESTINATION", "DESTINO FINAL" },
            ["SALIDA TOPE"]   = new[] { "SALIDA TOPE", "TOPE SALIDA", "SALIDA_TOPE", "TOPE", "TOPE DE SALIDA" },
            ["PRECINTO"]      = new[] { "PRECINTO", "SEAL", "SEALS", "Nº PRECINTO", "NUM PRECINTO", "NUMERO PRECINTO" }
        };

        private static string N(string s) =>
            (s ?? string.Empty).Trim().ToUpperInvariant()
                .Replace("Á","A").Replace("É","E").Replace("Í","I")
                .Replace("Ó","O").Replace("Ú","U").Replace("Ü","U").Replace("Ñ","N");

        public static List<LoadRow> ImportExcel(string path)
        {
            using var wb = new XLWorkbook(path);

            // 1) Elegir la HOJA que mejor coincide con nuestros encabezados
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

            // Si no hay hoja con al menos los 3 básicos, devolvemos vacío (UI avisará)
            if (bestWs == null || !HasMinimum(bestMap))
                return new List<LoadRow>();

            // 2) Volcar datos desde esa hoja
            var wsBest = bestWs;
            var lastRow = wsBest.LastRowUsed()?.RowNumber() ?? 0;
            var lastCol = wsBest.LastColumnUsed()?.ColumnNumber() ?? 0;

            var list = new List<LoadRow>();

            for (int r = bestHeaderRow + 1; r <= lastRow; r++)
            {
                bool empty = true;
                for (int c = 1; c <= lastCol; c++)
                {
                    if (!string.IsNullOrWhiteSpace(wsBest.Cell(r, c).GetString()))
                    {
                        empty = false; break;
                    }
                }
                if (empty) continue;

                string GetText(string key)
                {
                    if (!bestMap.TryGetValue(key, out var c)) return "";
                    return wsBest.Cell(r, c).GetString();
                }

                TimeSpan? GetTime(string key)
                {
                    if (!bestMap.TryGetValue(key, out var c)) return null;
                    var cell = wsBest.Cell(r, c);

                    if (cell.DataType == XLDataType.Number)
                    {
                        // Excel: días decimales
                        var d = cell.GetDouble();
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

                    // normalizar separadores
                    s = s.Replace(".", ":").Replace(",", ":");
                    if (TimeSpan.TryParse(s, CultureInfo.InvariantCulture, out var t1))
                        return new TimeSpan(t1.Hours, t1.Minutes, 0);

                    // “730” → “07:30”
                    var onlyDigits = new string(s.Where(char.IsDigit).ToArray());
                    if (onlyDigits.Length is 3 or 4)
                    {
                        onlyDigits = onlyDigits.PadLeft(4,'0');
                        var fixedStr = onlyDigits.Insert(2, ":");
                        if (TimeSpan.TryParse(fixedStr, CultureInfo.InvariantCulture, out var t2))
                            return new TimeSpan(t2.Hours, t2.Minutes, 0);
                    }

                    return null;
                    }

                var row = new LoadRow
                {
                    Transportista = GetText("TRANSPORTISTA"),
                    Matricula     = GetText("MATRICULA"),
                    Destino       = GetText("DESTINO"),
                    Muelle        = GetText("MUELLE"),
                    Estado        = GetText("ESTADO"),
                    Precinto      = GetText("PRECINTO"),
                    SalidaTope    = GetTime("SALIDA TOPE")
                };

                list.Add(row);
            }

            // 3) Orden lógico por muelle (numérico si hay dígitos)
            var ordered = list
                .OrderBy(x =>
                {
                    var digits = new string((x.Muelle ?? "").Where(char.IsDigit).ToArray());
                    return int.TryParse(digits, out var n) ? n : int.MaxValue;
                })
                .ThenBy(x => x.Muelle)
                .ToList();

            return ordered;
        }

        // Detecta fila de encabezados y columnas; devuelve (#aciertos, fila, mapa)
        private static (int hits, int headerRow, Dictionary<string,int> map) DetectHeaders(IXLWorksheet ws)
        {
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;

            int bestHits = -1;
            int headerRow = 1;
            var bestMap = new Dictionary<string,int>();

            // miramos primeras 100 filas o hasta la última usada
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
            // Mínimo imprescindible para trabajar
            string[] needed = { "MATRICULA", "DESTINO", "MUELLE", "ESTADO" };
            return needed.All(map.ContainsKey);
        }
    }
}
