using System;
using System.ComponentModel;

namespace Plmeco.App.Models
{
    public class LoadRow : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        private int _id;
        public int Id { get => _id; set { _id = value; OnPropertyChanged(nameof(Id)); } }

        private string _transportista = "";
        public string Transportista { get => _transportista; set { _transportista = value; OnPropertyChanged(nameof(Transportista)); } }

        private string _matricula = "";
        public string Matricula { get => _matricula; set { _matricula = value; OnPropertyChanged(nameof(Matricula)); } }

        private string _destino = "";
        public string Destino { get => _destino; set { _destino = value; OnPropertyChanged(nameof(Destino)); } }

        private string _muelle = "";
        public string Muelle { get => _muelle; set { _muelle = value; OnPropertyChanged(nameof(Muelle)); } }

        private string _estado = "";
        public string Estado
        {
            get => _estado;
            set { _estado = value; OnPropertyChanged(nameof(Estado)); OnPropertyChanged(nameof(Incidencias)); }
        }

        // 🔹 NUEVO: PRECINTO (editable)
        private string _precinto = "";
        public string Precinto
        {
            get => _precinto;
            set { _precinto = value; OnPropertyChanged(nameof(Precinto)); }
        }

        private TimeSpan? _llegadaReal;
        public TimeSpan? LlegadaReal
        {
            get => _llegadaReal;
            set { _llegadaReal = value; OnPropertyChanged(nameof(LlegadaReal)); OnPropertyChanged(nameof(Incidencias)); }
        }

        private TimeSpan? _salidaReal;
        public TimeSpan? SalidaReal
        {
            get => _salidaReal;
            set { _salidaReal = value; OnPropertyChanged(nameof(SalidaReal)); OnPropertyChanged(nameof(Incidencias)); }
        }

        private TimeSpan? _salidaTope;
        public TimeSpan? SalidaTope
        {
            get => _salidaTope;
            set { _salidaTope = value; OnPropertyChanged(nameof(SalidaTope)); OnPropertyChanged(nameof(Incidencias)); }
        }

        private bool _lex;
        public bool Lex { get => _lex; set { _lex = value; OnPropertyChanged(nameof(Lex)); } }

        // Incidencias calculadas automáticamente
        public string Incidencias =>
            (SalidaReal.HasValue && SalidaTope.HasValue && SalidaReal > SalidaTope)
            ? "RETRASO TRANSPORTISTA"
            : string.Empty;
    }
}
