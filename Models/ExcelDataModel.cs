using System.ComponentModel;

namespace ExcelDynamicViewer.Models;

public class ExcelDataModel : INotifyPropertyChanged
{
    private Dictionary<string, object?> _data = new();
    public Dictionary<string, object?> Data
    {
        get => _data;
        set
        {
            _data = value;
            OnPropertyChanged(nameof(Data));
        }
    }

    private int _index;
    public int Index
    {
        get => _index;
        set
        {
            _index = value;
            OnPropertyChanged(nameof(Index));
        }
    }

    public object? this[string columnName]
    {
        get => _data.ContainsKey(columnName) ? _data[columnName] : null;
        set
        {
            _data[columnName] = value;
            OnPropertyChanged($"Item[{columnName}]");
        }
    }


    public event PropertyChangedEventHandler? PropertyChanged;
    protected virtual void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}