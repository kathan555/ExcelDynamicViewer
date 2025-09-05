using System.Data;
using System.IO;
using System.Text;
using ExcelDataReader;
using ExcelDynamicViewer.Models;

namespace ExcelDynamicViewer.Services;

public class ExcelService : IExcelService
{
    /// <summary>
    /// read Excel file and return structured data.
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    public ExcelDataResult ReadExcel(string filePath)
    {
        var result = new ExcelDataResult();

        // Required for ExcelDataReader to handle encoding properly
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
            {
                UseHeaderRow = true
            }
        });

        if (dataSet.Tables.Count == 0)
            return result;

        DataTable dataTable = dataSet.Tables[0];

        // Get column headers
        for (int col = 0; col < dataTable.Columns.Count; col++)
        {
            string header = dataTable.Columns[col].ColumnName;
            if (string.IsNullOrEmpty(header))
                header = $"Column {col + 1}";

            Type dataType = DetectDataType(dataTable, col);

            result.Columns.Add(new ExcelColumn
            {
                Header = header,
                DataType = dataType,
                Index = col
            });
        }

        // Get data rows
        for (int row = 0; row < dataTable.Rows.Count; row++)
        {
            ExcelDataModel dataModel = new ExcelDataModel();
            Dictionary<string, object?> dataDict = new ();

            for (int col = 0; col < result.Columns.Count; col++)
            {
                var column = result.Columns[col];
                var cellValue = dataTable.Rows[row][col];

                if (column.DataType != null)
                {
                    object? convertedValue = ConvertValue(cellValue, column.DataType);
                    if (column.Header != null && convertedValue != null) 
                        dataDict[column.Header] = convertedValue;
                }
            }

            dataModel.Data = dataDict;
            result.Data.Add(dataModel);
        }

        return result;
    }

    /// <summary>
    /// Detects the data type of column by inspecting its values.
    /// </summary>
    /// <param name="dataTable"></param>
    /// <param name="columnIndex"></param>
    /// <returns></returns>
    private Type DetectDataType(DataTable dataTable, int columnIndex)
    {
        // Check first 10 rows to determine data type
        for (int row = 0; row < Math.Min(10, dataTable.Rows.Count); row++)
        {
            var cellValue = dataTable.Rows[row][columnIndex];

            if (cellValue is DateTime)
                return typeof(DateTime);

            if (cellValue is double || cellValue is decimal || cellValue is int || cellValue is float)
                return typeof(double);

            if (cellValue is bool)
                return typeof(bool);

            if (cellValue is string strValue)
            {
                // Try to parse as number
                if (double.TryParse(strValue, out _))
                    return typeof(double);

                // Try to parse as date
                if (DateTime.TryParse(strValue, out _))
                    return typeof(DateTime);

                // Try to parse as boolean
                if (bool.TryParse(strValue, out _))
                    return typeof(bool);
            }
        }

        return typeof(string);
    }

    /// <summary>
    /// Converts a cell value to the specified target type.
    /// </summary>
    /// <param name="value"></param>
    /// <param name="targetType"></param>
    /// <returns></returns>
    private object? ConvertValue(object? value, Type targetType)
    {
        if (value == null)
            return targetType.IsValueType ? Activator.CreateInstance(targetType) : null;

        if (targetType == typeof(string))
            return value.ToString();

        if (targetType == typeof(DateTime))
        {
            if (value is DateTime dt)
                return dt;

            if (DateTime.TryParse(value.ToString(), out var parsedDate))
                return parsedDate;

            return null;
        }

        if (targetType == typeof(double))
        {
            if (value is double d)
                return d;

            if (double.TryParse(value.ToString(), out var parsedDouble))
                return parsedDouble;

            return 0.0;
        }

        if (targetType == typeof(bool))
        {
            if (value is bool b)
                return b;

            if (bool.TryParse(value.ToString(), out var parsedBool))
                return parsedBool;

            return false;
        }

        return value;
    }
}