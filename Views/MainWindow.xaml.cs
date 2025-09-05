using System.Windows.Controls;
using ExcelDynamicViewer.ViewModels;

namespace ExcelDynamicViewer.Views;

public partial class MainWindow 
{

    public MainWindow()
    {
        InitializeComponent();
        DataContext = new MainViewModel(GenerateDataGridColumns);
    }


    /// <summary>
    /// Dynamically generate DataGrid columns based on the loaded Excel data
    /// </summary>
    private void GenerateDataGridColumns()
    {
        if (DataContext is MainViewModel viewModel)
        {
            // Clear existing columns
            ExcelDataGrid.Columns.Clear();

            if (viewModel.Columns.Count == 0)
                return;
            
            // Add row number column
            ExcelDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "#",
                Width = new DataGridLength(1, DataGridLengthUnitType.Auto),
                Binding = new System.Windows.Data.Binding("Index"),
                IsReadOnly = true
            });

            // Add dynamic columns from Excel
            foreach (var column in viewModel.Columns)
            {
                ExcelDataGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = column.Header,
                    Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                    Binding = new System.Windows.Data.Binding($"Data[{column.Header}]"),
                    IsReadOnly = true
                });
            }
        }
    }

    /// <summary>
    /// Prevent auto-generation of columns since we're creating them manually
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void ExcelDataGrid_OnAutoGeneratingColumn(object? sender, DataGridAutoGeneratingColumnEventArgs e)
    {
        // Cancel auto-generation since we're creating columns manually
        e.Cancel = true;
    }
}