using System.Data;
using System.Drawing;
using BalaReva.EasyExcel.Utilities;
using Interop = Microsoft.Office.Interop.Excel;

namespace BalaReva.EasyExcel;

public sealed partial class ExcelWorkbook
{
    /// <summary>The table on a sheet, by name when given and by position otherwise.</summary>
    private Interop.ListObject TableOn(string sheetName, TableRef table)
    {
        ArgumentNullException.ThrowIfNull(table);
        var tables = Sheet(sheetName).ListObjects;

        if (!string.IsNullOrWhiteSpace(table.TableName))
        {
            foreach (Interop.ListObject candidate in tables)
                if (string.Equals(candidate.Name, table.TableName, StringComparison.OrdinalIgnoreCase))
                    return candidate;

            throw new ArgumentException($"No table called '{table.TableName}'.", nameof(table));
        }

        var index = table.TableIndex < 1 ? 1 : table.TableIndex;
        if (index > tables.Count)
            throw new ArgumentOutOfRangeException(
                nameof(table), index, $"The sheet has {tables.Count} table(s).");

        return tables[index];
    }

    /// <inheritdoc />
    public void DeleteTable(string sheetName, TableRef table)
    {
        TableOn(sheetName, table).Unlist();
        Save();
    }

    /// <inheritdoc />
    public void ResizeTable(string sheetName, TableRef table, string newRange)
    {
        var target = TableOn(sheetName, table);
        target.Resize(Sheet(sheetName).Range[newRange]);
        Save();
    }

    /// <inheritdoc />
    public bool TableExists(string sheetName, string tableName)
    {
        foreach (Interop.ListObject table in Sheet(sheetName).ListObjects)
            if (string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase))
                return true;
        return false;
    }

    /// <inheritdoc />
    public (List<string> Names, DataTable Table) GetTableNames(string sheetName)
    {
        var names = new List<string>();
        var table = new DataTable("Tables");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Range", typeof(string));

        foreach (Interop.ListObject listObject in Sheet(sheetName).ListObjects)
        {
            names.Add(listObject.Name);
            table.Rows.Add(listObject.Name, listObject.Range.Address[false, false]);
        }
        return (names, table);
    }

    /// <inheritdoc />
    public void RemoveFilter(string sheetName, int tableIndex)
    {
        var table = TableOn(sheetName, new TableRef { TableIndex = tableIndex });
        if (table.AutoFilter is not null) table.AutoFilter.ShowAllData();
        Save();
    }

    /// <inheritdoc />
    public void RemoveFilterNonTable(string sheetName)
    {
        var sheet = Sheet(sheetName);
        if (sheet.AutoFilterMode) sheet.ShowAllData();
        Save();
    }

    /// <inheritdoc />
    public void Sorting(string sheetName, TableRef table, string[] columnNames,
                        int[] columnIndexes, SortingTableEnum order)
    {
        var target = TableOn(sheetName, table);
        var sort = target.Sort;
        sort.SortFields.Clear();

        var direction = order == SortingTableEnum.Ascending
            ? Interop.XlSortOrder.xlAscending
            : Interop.XlSortOrder.xlDescending;

        foreach (var name in columnNames ?? [])
            sort.SortFields.Add(target.ListColumns[name].Range, Order: direction);

        foreach (var index in columnIndexes ?? [])
            sort.SortFields.Add(target.ListColumns[index].Range, Order: direction);

        if (sort.SortFields.Count == 0)
            throw new ArgumentException(
                "Sorting needs at least one column, by name or by index.", nameof(columnNames));

        sort.Apply();
        Save();
    }

    // ---------------------------------------------------------------- workbook

    /// <inheritdoc />
    public void ExportToTextFile(string sheetName, string textFilePath)
    {
        var sheet = Sheet(sheetName);
        sheet.Copy();

        var copy = _application.ActiveWorkbook;
        try
        {
            copy.SaveAs(textFilePath, Interop.XlFileFormat.xlTextWindows);
        }
        finally
        {
            copy.Close(SaveChanges: false);
        }
    }

    /// <inheritdoc />
    public void MergeSheetByRow(string sheetName, MergeRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        var destination = Sheet(sheetName);

        var source = _application.Workbooks.Open(
            Filename: request.AppendFileName,
            ReadOnly: true,
            Password: request.AppendFilePassword,
            WriteResPassword: request.AppendModifyPassword);

        try
        {
            var sourceSheet = string.IsNullOrWhiteSpace(request.AppendSheet)
                ? (Interop.Worksheet)source.ActiveSheet
                : (Interop.Worksheet)source.Worksheets[request.AppendSheet];

            var range = string.IsNullOrWhiteSpace(request.AppendCellRange)
                ? sourceSheet.UsedRange
                : sourceSheet.Range[request.AppendCellRange];

            var used = destination.UsedRange;
            var firstFreeRow = used.Row + used.Rows.Count;
            var column = string.IsNullOrWhiteSpace(request.StartColumn) ? "A" : request.StartColumn;

            range.Copy(destination.Range[$"{column}{firstFreeRow}"]);
        }
        finally
        {
            source.Close(SaveChanges: false);
        }
        Save();
    }

    /// <inheritdoc />
    public void MoveSheet(string sheetName, string beforeSheet)
    {
        var sheet = Sheet(sheetName);

        if (string.IsNullOrWhiteSpace(beforeSheet))
            sheet.Move(After: _workbook.Worksheets[_workbook.Worksheets.Count]);
        else
            sheet.Move(Before: Sheet(beforeSheet));

        Save();
    }

    /// <inheritdoc />
    public void SaveAsWorkBook(string fileName, FileFormatEnum format, string password, string writeResPassword) =>
        _workbook.SaveAs(
            Filename: fileName,
            FileFormat: (Interop.XlFileFormat)(int)format,
            Password: password,
            WriteResPassword: writeResPassword);

    /// <inheritdoc />
    public string[] ShowVisibleSheet(VisibleInvisibleEnum visibleType)
    {
        var wanted = visibleType == VisibleInvisibleEnum.Visible;
        var sheets = new List<string>();

        foreach (Interop.Worksheet sheet in _workbook.Worksheets)
        {
            var visible = sheet.Visible == Interop.XlSheetVisibility.xlSheetVisible;
            if (visible == wanted) sheets.Add(sheet.Name);
        }
        return [.. sheets];
    }

    /// <inheritdoc />
    public void TabColor(string sheetName, Color color)
    {
        var sheet = Sheet(sheetName);

        if (color.IsEmpty) sheet.Tab.ColorIndex = Interop.XlColorIndex.xlColorIndexNone;
        else sheet.Tab.Color = Bgr(color);

        Save();
    }
}
