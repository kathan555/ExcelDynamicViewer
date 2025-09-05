namespace ExcelDynamicViewer.Models;

public class ExcelDataResult
{
    public List<ExcelColumn> Columns { get; set; } = new();
    public List<ExcelDataModel> Data { get; set; } = new();
}