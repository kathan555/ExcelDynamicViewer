namespace ExcelDynamicViewer.Models;

public class ExcelColumn
{
    public string? Header { get; set; }
    public Type? DataType { get; set; }
    public int Index { get; set; }
}