using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Plmeco.App.Models;

namespace Plmeco.App.Services
{
    public static class PersistenceService
    {
        private static readonly string BaseDir =
            Environment.GetEnvironmentVariable("PLMECO_DATA_DIR")
            ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PLMECO");

        private static readonly string FilePath = Path.Combine(BaseDir, "autosave_multi.json");

        private static readonly JsonSerializerOptions JsonOpts = new()
        {
            WriteIndented = true,
        };

        public class DocumentSnapshot
        {
            public string Title { get; set; } = "Sin título";
            public string? CurrentFile { get; set; }
            public List<LoadRow> Rows { get; set; } = new();
        }

        public class AppSnapshot
        {
            public List<DocumentSnapshot> Documents { get; set; } = new();
            public int SelectedIndex { get; set; }
        }

        private static void EnsureDir()
        {
            if (!Directory.Exists(BaseDir)) Directory.CreateDirectory(BaseDir);
        }

        public static void Save(List<DocumentSnapshot> docs, int selectedIndex)
        {
            try
            {
                EnsureDir();
                var snapshot = new AppSnapshot { Documents = docs, SelectedIndex = selectedIndex };
                File.WriteAllText(FilePath, JsonSerializer.Serialize(snapshot, JsonOpts));
            }
            catch { /* silencioso */ }
        }

        public static AppSnapshot Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new AppSnapshot();
                var txt = File.ReadAllText(FilePath);
                var snap = JsonSerializer.Deserialize<AppSnapshot>(txt, JsonOpts);
                return snap ?? new AppSnapshot();
            }
            catch
            {
                return new AppSnapshot();
            }
        }

        public static void Clear()
        {
            try { if (File.Exists(FilePath)) File.Delete(FilePath); } catch { }
        }
    }
}
