using ExcelDynamicViewer.Models;

namespace ExcelDynamicViewer.Services;

public interface IExcelService
{
    ExcelDataResult ReadExcel(string filePath);
}