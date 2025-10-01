using System.Collections.ObjectModel;
using System.ComponentModel;

namespace Plmeco.App.Models
{
    public class DocumentView : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string n) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));

        private string _title = "Sin título";
        public string Title { get => _title; set { _title = value; OnPropertyChanged(nameof(Title)); } }

        private string? _currentFile;
        public string? CurrentFile { get => _currentFile; set { _currentFile = value; OnPropertyChanged(nameof(CurrentFile)); } }

        public ObservableCollection<LoadRow> Rows { get; } = new();
    }
}
