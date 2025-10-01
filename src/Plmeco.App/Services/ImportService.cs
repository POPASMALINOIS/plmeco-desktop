using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using Plmeco.App.Models;

namespace Plmeco.App.Services
{
    public static class ImportService
    {
        // Sinónimos por columna (insensibles a acentos/mayúsculas)
        private static readonly Dictionary<string, string[]> Syns = new()
        {
            ["TRANSPORTISTA"] = new[] { "TRANSPORTISTA", "TRANSPORTE", "CARRIER", "TRANSPORTER" },
            ["MATRICULA"]     = new[] { "MATRICULA", "MATRÍCULA", "PLACA", "MATRIC", "REG", "REGISTRO" },
            ["MUELLE"]        = new[] { "MUELLE", "MUEL.", "DOCK" },
            ["ESTADO"]        = new[] { "ESTADO", "STATUS", "EST." },
            ["DESTINO"]       = new[] { "DESTINO", "DEST.", "DESTINATION", "DESTINO FINAL" },
            ["SALIDA TOPE"]   = new[] { "SALIDA TOPE", "TOPE SALIDA", "SALIDA_TOPE", "TOPE" },
            // 🔹 NUEVO: PRECINTO
            ["PRECINTO"]      = new[] { "PRECINTO", "SEAL", "SEALS", "Nº PRECINTO", "NUM PRECINTO", "NUMERO PRECINTO" }
        };

        /// <summary>
        /// Importa la hoja Excel y devuelve la lista lista para bindear al DataGrid.
        /// </summary>
        public static List<LoadRow> ImportExcel(string path)
        {
            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheets.First();

            var lastRow = ws.LastRowUsed().RowNumber();
            var lastCol = ws.LastColumnUsed().ColumnNumber();

            // 1) Localizar la fila de encabezados que mejor encaja con nuestros sinónimos
            int headerRow = 1, bestHits = -1;
            Dictionary<string, int> colMap = new();

            for (int r = 1; r <= Math.Min(100, lastRow); r++)
            {
                var temp = new Dictionary<string, int>();
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

            // 2) Funciones de ayuda para leer valores
            string GetText(int row, string key)
            {
                if (!colMap.TryGetValue(key, out var c)) return "";
                return ws.Cell(row, c).GetString();
            }

            TimeSpan? GetTime(int row, string key)
            {
                if (!colMap.TryGetValue(key, out var c)) return null;
                var cell = ws.Cell(row, c);

                // Excel puede guardar horas como número (días decimales) o DateTime
                if (cell.DataType == XLDataType.Number)
                {
                    var d = cell.GetDouble();               // días
                    var ts = TimeSpan.FromDays(d);
                    return new TimeSpan(ts.Hours, ts.Minutes, 0);
                }
                if (cell.DataType == XLDataType.DateTime)
                {
                    var dt = cell.GetDateTime();
                    return new TimeSpan(dt.Hour, dt.Minute, 0);
                }

                var s = cell.GetString();
                if (string.IsNullOrWhiteSpace(s)) return null;

                if (TimeSpan.TryParse(s, out var t1))
                    return new TimeSpan(t1.Hours, t1.Minutes, 0);

                // Reparaciones típicas ("730" -> "07:30", "7.30" -> "7:30")
                var cleaned = s.Replace(".", ":").Replace(",", ":").Trim();
                if (cleaned.Length is 3 or 4 && cleaned.All(char.IsDigit))
                {
                    cleaned = cleaned.PadLeft(4, '0');
                    cleaned = cleaned.Insert(2, ":");
                }
                if (TimeSpan.TryParse(cleaned, out var t2))
                    return new TimeSpan(t2.Hours, t2.Minutes, 0);

                return null;
            }

            // 3) Volcar filas
            var list = new List<LoadRow>();
            for (int r = headerRow + 1; r <= lastRow; r++)
            {
                // Saltar filas completamente vacías
                bool empty = true;
                for (int c = 1; c <= lastCol; c++)
                {
                    if (!string.IsNullOrWhiteSpace(ws.Cell(r, c).GetString()))
                    {
                        empty = false; break;
                    }
                }
                if (empty) continue;

                var row = new LoadRow
                {
                    Transportista = GetText(r, "TRANSPORTISTA"),
                    Matricula     = GetText(r, "MATRICULA"),
                    Destino       = GetText(r, "DESTINO"),
                    Muelle        = GetText(r, "MUELLE"),
                    Estado        = GetText(r, "ESTADO"),
                    Precinto      = GetText(r, "PRECINTO"),    // 🔹 NUEVO
                    SalidaTope    = GetTime (r, "SALIDA TOPE")
                };

                list.Add(row);
            }

            // 4) Orden “natural” por muelle (si tu muelle son números, intenta ordenar numérico)
            // Primero probamos a extraer número; si no, usamos orden alfabético
            var ordered = list
                .OrderBy(x =>
                {
                    if (int.TryParse(new string((x.Muelle ?? "").Where(char.IsDigit).ToArray()), out var num))
                        return num;
                    return int.MaxValue;
                })
                .ThenBy(x => x.Muelle)
                .ToList();

            return ordered;
        }

        // Normalización sencilla para comparar encabezados
        private static string Normalize(string s) =>
            (s ?? string.Empty).Trim().ToUpperInvariant()
              .Replace("Á", "A").Replace("É", "E").Replace("Í", "I")
              .Replace("Ó", "O").Replace("Ú", "U").Replace("Ü", "U")
              .Replace("Ñ", "N");
    }
}
