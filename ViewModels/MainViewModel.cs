using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using CommunityToolkit.Mvvm.Input;
using ExcelDynamicViewer.Infrastructure;
using ExcelDynamicViewer.Models;
using ExcelDynamicViewer.Services;
using Microsoft.Win32;

namespace ExcelDynamicViewer.ViewModels;

internal class MainViewModel : ViewModelBase
{
    private readonly IExcelService _excelService;
    private readonly Action _generateDataGridColumns;


    private ObservableCollection<ExcelDataModel> _excelData = new ();
    public ObservableCollection<ExcelDataModel> ExcelData
    {
        get => _excelData;
        set
        {
            _excelData = value;
            OnPropertyChanged();
        }
    }

    private ObservableCollection<ExcelColumn> _columns = new ();
    public ObservableCollection<ExcelColumn> Columns
    {
        get => _columns;
        set
        {
            _columns = value;
            OnPropertyChanged();
        }
    }

    private string _statusMessage = string.Empty;
    public string StatusMessage
    {
        get => _statusMessage;
        set
        {
            _statusMessage = value;
            OnPropertyChanged();
        }
    }


    public ICommand LoadExcelCommand => new RelayCommand(LoadExcel);


    public MainViewModel(Action generateDataGridColumns)
    {
        _generateDataGridColumns = generateDataGridColumns;
        _excelService = new ExcelService();
        StatusMessage = "Ready to load Excel file";
    }


    /// <summary>
    /// Load Excel file.
    /// </summary>
    private void LoadExcel()
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "Excel Files|*.xlsx;*.xls;*.csv",
            Title = "Select Excel File"
        };

        if (openFileDialog.ShowDialog() == true)
        {
            LoadExcelDataAsync(openFileDialog.FileName);
        }
    }

    /// <summary>
    /// Load Excel data in a background task to keep UI responsive
    /// </summary>
    /// <param name="filePath"></param>
    private void LoadExcelDataAsync(string filePath)
    {
        StatusMessage = "Loading Excel data...";

        try
        {
            ExcelData.Clear();
            Columns.Clear();

            ExcelDataResult result = Task.Run(() => _excelService.ReadExcel(filePath)).WaitAsync(CancellationToken.None).Result;

            foreach (var column in result.Columns)
            {
                Columns.Add(column);
            }

            int index = 1;
            foreach (var data in result.Data)
            {
                // Add index property for row numbering
                data.Index = index++;
                ExcelData.Add(data);
            }

            StatusMessage = $"Loaded {ExcelData.Count - 1} rows with {result.Columns.Count} columns from {Path.GetFileName(filePath)}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading Excel file: {ex.Message}", "Error",
                MessageBoxButton.OK, MessageBoxImage.Error);
            StatusMessage = "Error loading Excel file";
        }
        finally
        {
            //Add refresh Ui here if needed
            _generateDataGridColumns.Invoke();
        }
    }
}