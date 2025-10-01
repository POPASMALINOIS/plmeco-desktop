using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using Plmeco.App.Models;

namespace Plmeco.App.Services
{
    public static class PersistenceService
    {
        // Puedes forzar una carpeta de red con la variable de entorno PLMECO_DATA_DIR
        private static readonly string BaseDir =
            Environment.GetEnvironmentVariable("PLMECO_DATA_DIR")
            ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PLMECO");

        private static readonly string FilePath = Path.Combine(BaseDir, "autosave.json");

        private static readonly JsonSerializerOptions JsonOpts = new()
        {
            WriteIndented = true,
            // En .NET 7/8 TimeSpan se soporta por defecto; añadimos este converter por compatibilidad
            Converters = { new JsonStringEnumConverter() }
        };

        public static void EnsureDir()
        {
            if (!Directory.Exists(BaseDir)) Directory.CreateDirectory(BaseDir);
        }

        public static void Save(IEnumerable<LoadRow> rows)
        {
            try
            {
                EnsureDir();
                var payload = JsonSerializer.Serialize(rows, JsonOpts);
                File.WriteAllText(FilePath, payload);
            }
            catch
            {
                // Silencioso: no queremos cortar la operativa por fallo de disco
            }
        }

        public static List<LoadRow> Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new List<LoadRow>();
                var json = File.ReadAllText(FilePath);
                var data = JsonSerializer.Deserialize<List<LoadRow>>(json, JsonOpts);
                return data ?? new List<LoadRow>();
            }
            catch
            {
                return new List<LoadRow>();
            }
        }

        public static void Clear()
        {
            try
            {
                if (File.Exists(FilePath)) File.Delete(FilePath);
            }
            catch { /* ignore */ }
        }
    }
}
