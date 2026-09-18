using System.Data;
using System.Runtime.InteropServices;
using BalaReva.Excel.Base;
using BalaReva.Excel.Enums;
using BalaReva.Excel.Sheets;
using BalaReva.Excel.Utilities;
using Interop = Microsoft.Office.Interop.Excel;

namespace BalaReva.Excel;

/// <summary>
/// <see cref="IExcelService"/> on top of Excel's COM object model.
/// </summary>
/// <remarks>
/// Requires Windows with Excel installed, so this type is never exercised by CI. Every
/// method is a mechanical translation; the judgement lives in the activities, where a
/// stand-in can cover it.
/// </remarks>
public sealed class ExcelService : IExcelService
{
    /// <summary>The shared instance used when a workflow registers no extension.</summary>
    public static ExcelService Instance { get; } = new();

    /// <inheritdoc />
    public IExcelWorkbook Open(string filePath, string filePassword, string modifyPassword) =>
        new ExcelWorkbook(filePath, filePassword, modifyPassword);

    /// <inheritdoc />
    public DataTable ClipboardToDataTable(bool hasHeader)
    {
        // Goes through Excel rather than the Windows clipboard API so that the package
        // needs no WinForms or WPF dependency for one activity. Excel pastes into a
        // scratch workbook, which is then read and discarded.
        var application = new Interop.Application { Visible = false, DisplayAlerts = false };
        try
        {
            var workbook = application.Workbooks.Add();
            var sheet = (Interop.Worksheet)workbook.Worksheets[1];
            sheet.Paste();

            var used = sheet.UsedRange;
            var table = ExcelWorkbook.ToDataTable(used, hasHeader);
            workbook.Close(SaveChanges: false);
            return table;
        }
        finally
        {
            try { application.Quit(); } catch (COMException) { /* already gone */ }
        }
    }
}

/// <summary>A workbook held open through COM.</summary>
internal sealed class ExcelWorkbook : IExcelWorkbook
{
    private readonly Interop.Application _application;
    private readonly Interop.Workbook _workbook;
    private bool _closed;

    internal ExcelWorkbook(string filePath, string filePassword, string modifyPassword)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("FilePath is required.", nameof(filePath));

        FilePath = filePath;
        _application = new Interop.Application { Visible = false, DisplayAlerts = false };
        try
        {
            _workbook = File.Exists(filePath)
                ? _application.Workbooks.Open(filePath,
                    Password: filePassword ?? string.Empty,
                    WriteResPassword: modifyPassword ?? string.Empty)
                : _application.Workbooks.Add();
        }
        catch
        {
            Quit();
            throw;
        }
    }

    private string FilePath { get; }

    // ---------------------------------------------------------------- charts

    /// <inheritdoc />
    public void DrawChart(string sheetName, ChartRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        var sheet = Sheet(sheetName);

        var charts = (Interop.ChartObjects)sheet.ChartObjects();
        var chartObject = charts.Add(
            request.Size.Left, request.Size.Top,
            request.Size.Width > 0 ? request.Size.Width : 360,
            request.Size.Height > 0 ? request.Size.Height : 240);

        var chart = chartObject.Chart;
        chart.SetSourceData(sheet.Range[request.CellRange]);
        chart.ChartType = (Interop.XlChartType)request.ChartType;

        if (!string.IsNullOrWhiteSpace(request.LegendRange))
        {
            var series = (Interop.Series)chart.SeriesCollection(1);
            series.XValues = sheet.Range[request.LegendRange];
        }

        if (!string.IsNullOrWhiteSpace(request.ChartTitle))
        {
            chart.HasTitle = true;
            chart.ChartTitle.Text = request.ChartTitle;
        }

        chart.HasLegend = request.Options.ShowLegend;
        chart.ApplyDataLabels(
            (Interop.XlDataLabelsType)(int)request.Options.DataLabelsType,
            LegendKey: request.ShowLegendKey,
            AutoText: request.Options.AutoText,
            HasLeaderLines: request.Options.HasLeaderLines,
            ShowSeriesName: request.Options.ShowSeriesName,
            ShowCategoryName: request.Options.ShowCategoryName,
            ShowValue: request.Options.ShowValue,
            ShowPercentage: request.Options.ShowPercentage,
            ShowBubbleSize: request.Options.ShowBubbleSize,
            Separator: request.Options.Separator);

        if (!string.IsNullOrWhiteSpace(request.ImageCopy))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(request.ImageCopy))!);
            chart.Export(request.ImageCopy);
        }

        Save();
    }

    // -------------------------------------------------------------- comments

    /// <inheritdoc />
    public void AddComment(string sheetName, string cell, string comment)
    {
        var range = Cell(sheetName, cell);
        // AddComment throws when one is already there, so replace rather than fail.
        range.ClearComments();
        range.AddComment(comment);
        Save();
    }

    /// <inheritdoc />
    public void DeleteComment(string sheetName, string cell)
    {
        Cell(sheetName, cell).ClearComments();
        Save();
    }

    /// <inheritdoc />
    public string GetComment(string sheetName, string cell)
    {
        var comment = Cell(sheetName, cell).Comment;
        return comment is null ? string.Empty : comment.Text();
    }

    /// <inheritdoc />
    public void ShowHideComment(string sheetName, string cell, bool show)
    {
        var comment = Cell(sheetName, cell).Comment
            ?? throw new InvalidOperationException($"Cell '{cell}' has no comment.");
        comment.Visible = show;
        Save();
    }

    // ------------------------------------------------------------------ data

    /// <inheritdoc />
    public void CopyData(string sheetName, string copyRange) =>
        Sheet(sheetName).Range[copyRange].Copy();

    /// <inheritdoc />
    public void DeleteData(string sheetName, string deleteRange)
    {
        Sheet(sheetName).Range[deleteRange].Clear();
        Save();
    }

    /// <inheritdoc />
    public void FindReplace(string sheetName, string cellRange, string find, string replace,
                            FindReplaceEnum option)
    {
        if (string.IsNullOrEmpty(find))
            throw new ArgumentException("Find is required.", nameof(find));

        var target = string.IsNullOrWhiteSpace(cellRange)
            ? Sheet(sheetName).UsedRange
            : Sheet(sheetName).Range[cellRange];

        target.Replace(find, replace ?? string.Empty,
            option == FindReplaceEnum.Whole ? Interop.XlLookAt.xlWhole : Interop.XlLookAt.xlPart);
        Save();
    }

    /// <inheritdoc />
    public void RemoveDuplicates(string sheetName, string cellRange, object[] columns, bool hasHeader)
    {
        var target = string.IsNullOrWhiteSpace(cellRange)
            ? Sheet(sheetName).UsedRange
            : Sheet(sheetName).Range[cellRange];

        // No columns given means "every column", which is what Excel's own dialog does.
        var keys = columns is { Length: > 0 }
            ? columns
            : [.. Enumerable.Range(1, target.Columns.Count).Cast<object>()];

        target.RemoveDuplicates(keys,
            hasHeader ? Interop.XlYesNoGuess.xlYes : Interop.XlYesNoGuess.xlNo);
        Save();
    }

    // ---------------------------------------------------------------- layout

    /// <inheritdoc />
    public void HideColumns(string sheetName, string[] columnNames, bool hide)
    {
        var sheet = Sheet(sheetName);
        foreach (var name in columnNames ?? [])
        {
            if (string.IsNullOrWhiteSpace(name)) continue;
            ((Interop.Range)sheet.Columns[name]).EntireColumn.Hidden = hide;
        }
        Save();
    }

    /// <inheritdoc />
    public void HideRows(string sheetName, int[] rowNumbers, bool hide)
    {
        var sheet = Sheet(sheetName);
        foreach (var row in rowNumbers ?? [])
        {
            if (row < 1) continue;
            ((Interop.Range)sheet.Rows[row]).EntireRow.Hidden = hide;
        }
        Save();
    }

    /// <inheritdoc />
    public void AutoFitColumns(string sheetName, string[] columnsRange, bool autoFit, double columnWidth)
    {
        var sheet = Sheet(sheetName);
        var targets = columnsRange is { Length: > 0 }
            ? columnsRange.Where(c => !string.IsNullOrWhiteSpace(c))
                          .Select(c => ((Interop.Range)sheet.Columns[c]).EntireColumn)
            : [sheet.UsedRange.EntireColumn];

        foreach (var column in targets)
        {
            if (autoFit) column.AutoFit();
            else if (columnWidth > 0) column.ColumnWidth = columnWidth;
        }
        Save();
    }

    /// <inheritdoc />
    public void AutoFitRows(string sheetName, int[] rowsRange, bool autoFit, double rowHeight)
    {
        var sheet = Sheet(sheetName);
        var targets = rowsRange is { Length: > 0 }
            ? rowsRange.Where(r => r >= 1).Select(r => ((Interop.Range)sheet.Rows[r]).EntireRow)
            : [sheet.UsedRange.EntireRow];

        foreach (var row in targets)
        {
            if (autoFit) row.AutoFit();
            else if (rowHeight > 0) row.RowHeight = rowHeight;
        }
        Save();
    }

    /// <inheritdoc />
    public void MergeCells(string sheetName, string mergeRange, string cellText,
                           AlignmentEnum horizontal, AlignmentEnum vertical)
    {
        var range = Sheet(sheetName).Range[mergeRange];
        range.Merge();
        if (!string.IsNullOrEmpty(cellText)) range.Value2 = cellText;
        range.HorizontalAlignment = (int)horizontal;
        range.VerticalAlignment = (int)vertical;
        Save();
    }

    /// <inheritdoc />
    public void UnMergeCells(string sheetName, string unMergeRange,
                             AlignmentEnum horizontal, AlignmentEnum vertical)
    {
        var range = Sheet(sheetName).Range[unMergeRange];
        range.UnMerge();
        range.HorizontalAlignment = (int)horizontal;
        range.VerticalAlignment = (int)vertical;
        Save();
    }

    /// <inheritdoc />
    public void FormatCells(string sheetName, string[] cellRanges, CellFormatRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        var sheet = Sheet(sheetName);

        foreach (var name in cellRanges ?? [])
        {
            if (string.IsNullOrWhiteSpace(name)) continue;
            var range = sheet.Range[name];

            // Every one of these enums carries a Select member meaning "leave alone",
            // so each has to be checked rather than cast straight across.
            if (request.Horizontal != FormatHorizontalEnum.Select)
                range.HorizontalAlignment = (int)request.Horizontal;
            if (request.Verticle != FormatVerticleEnum.Select)
                range.VerticalAlignment = (int)request.Verticle;
            if (request.WrapText != FormatTextControlEnum.Select)
                range.WrapText = request.WrapText == FormatTextControlEnum.True;
            if (request.ShrinkFit != FormatTextControlEnum.Select)
                range.ShrinkToFit = request.ShrinkFit == FormatTextControlEnum.True;
            if (request.MergeCells != FormatTextControlEnum.Select)
                range.MergeCells = request.MergeCells == FormatTextControlEnum.True;
            if (request.TextDirection != FormatTextDirection.Select)
                range.ReadingOrder = (int)request.TextDirection;
            if (request.TextOrientation != TextOrientationEumn.Select)
                range.Orientation = (int)request.TextOrientation;
            else if (request.OrientationDeg != 0)
                range.Orientation = request.OrientationDeg;
            if (request.Indent > 0) range.IndentLevel = request.Indent;
        }
        Save();
    }

    /// <inheritdoc />
    public void ChangeCellType(string sheetName, string cell, string cellFormat)
    {
        Cell(sheetName, cell).NumberFormat = cellFormat;
        Save();
    }

    /// <inheritdoc />
    public void InsertTableFormat(string sheetName, string cellRange,
                                  TableFormatEnum style, string customStyle)
    {
        var range = string.IsNullOrWhiteSpace(cellRange)
            ? Sheet(sheetName).UsedRange
            : Sheet(sheetName).Range[cellRange];

        if (!string.IsNullOrWhiteSpace(customStyle))
        {
            var table = Sheet(sheetName).ListObjects.Add(
                Interop.XlListObjectSourceType.xlSrcRange, range, Type.Missing,
                Interop.XlYesNoGuess.xlYes);
            table.TableStyle = customStyle;
        }
        else
        {
            range.AutoFormat((Interop.XlRangeAutoFormat)(int)style);
        }
        Save();
    }

    /// <inheritdoc />
    public void SetTableFormat(string sheetName, int tableIndex,
                               TableFormatEnum style, string customStyle)
    {
        var tables = Sheet(sheetName).ListObjects;
        if (tableIndex < 1 || tableIndex > tables.Count)
        {
            throw new ArgumentOutOfRangeException(
                nameof(tableIndex), tableIndex, $"Must be between 1 and {tables.Count}.");
        }

        var table = tables[tableIndex];
        if (!string.IsNullOrWhiteSpace(customStyle)) table.TableStyle = customStyle;
        else table.Range.AutoFormat((Interop.XlRangeAutoFormat)(int)style);
        Save();
    }

    // ------------------------------------------------------- images and links

    /// <inheritdoc />
    public void InsertImage(string sheetName, string imagePath, IObjectSize size)
    {
        RequireFile(imagePath);
        ArgumentNullException.ThrowIfNull(size);

        // Pictures().Insert rather than Shapes.AddPicture: the latter takes an
        // MsoTriState, which lives in the separate office PIA, and pulling that in for
        // two calls is not worth an extra dependency.
        var picture = ((Interop.Pictures)Sheet(sheetName).Pictures()).Insert(imagePath);
        picture.Left = size.Left;
        picture.Top = size.Top;
        if (size.Width > 0) picture.Width = size.Width;
        if (size.Height > 0) picture.Height = size.Height;
        Save();
    }

    /// <inheritdoc />
    public void InsertImageAtCell(string sheetName, string cell, string imagePath,
                                  float imageWidth, float imageHeight)
    {
        RequireFile(imagePath);
        var anchor = Cell(sheetName, cell);

        var picture = ((Interop.Pictures)Sheet(sheetName).Pictures()).Insert(imagePath);
        picture.Left = (double)anchor.Left;
        picture.Top = (double)anchor.Top;
        picture.Width = imageWidth > 0 ? imageWidth : (double)anchor.Width;
        picture.Height = imageHeight > 0 ? imageHeight : (double)anchor.Height;
        Save();
    }

    /// <inheritdoc />
    public void ExtractGraphImage(string sheetName, string imageFolder, FileExtensEnum extension)
    {
        if (string.IsNullOrWhiteSpace(imageFolder))
            throw new ArgumentException("ImageFolder is required.", nameof(imageFolder));
        Directory.CreateDirectory(imageFolder);

        var charts = (Interop.ChartObjects)Sheet(sheetName).ChartObjects();
        for (var i = 1; i <= charts.Count; i++)
        {
            var chart = ((Interop.ChartObject)charts.Item(i)).Chart;
            chart.Export(Path.Combine(
                imageFolder, $"Chart{i}.{extension.ToString().ToLowerInvariant()}"));
        }
    }

    /// <inheritdoc />
    public void AddHyperlinks(string sheetName, DataTable inputTable, string[] columnNames, bool overwriteText)
    {
        ArgumentNullException.ThrowIfNull(inputTable);
        var sheet = Sheet(sheetName);

        foreach (var column in columnNames ?? [])
        {
            if (string.IsNullOrWhiteSpace(column)) continue;
            for (var row = 0; row < inputTable.Rows.Count; row++)
            {
                var address = inputTable.Rows[row][column]?.ToString();
                if (string.IsNullOrWhiteSpace(address)) continue;

                // Row 1 holds the header, so the first data row is row 2.
                var cell = (Interop.Range)sheet.Columns[column];
                var target = (Interop.Range)cell.Cells[row + 2, 1];
                sheet.Hyperlinks.Add(target, address,
                    TextToDisplay: overwriteText ? address : target.Text?.ToString());
            }
        }
        Save();
    }

    /// <inheritdoc />
    public void RemoveHyperlink(string sheetName, string cellRange)
    {
        var range = string.IsNullOrWhiteSpace(cellRange)
            ? Sheet(sheetName).UsedRange
            : Sheet(sheetName).Range[cellRange];
        range.Hyperlinks.Delete();
        Save();
    }

    // --------------------------------------------------- workbook and sheets

    /// <inheritdoc />
    public void AddSheet(string sheetName)
    {
        var sheet = (Interop.Worksheet)_workbook.Worksheets.Add(
            After: _workbook.Worksheets[_workbook.Worksheets.Count]);
        if (!string.IsNullOrWhiteSpace(sheetName)) sheet.Name = sheetName;
        Save();
    }

    /// <inheritdoc />
    public void DeleteSheet(string sheetName)
    {
        if (_workbook.Worksheets.Count == 1)
            throw new InvalidOperationException("A workbook must keep at least one sheet.");

        Sheet(sheetName).Delete();
        Save();
    }

    /// <inheritdoc />
    public void RenameSheet(string sheetName, string newSheetName)
    {
        if (string.IsNullOrWhiteSpace(newSheetName))
            throw new ArgumentException("NewSheetName is required.", nameof(newSheetName));

        Sheet(sheetName).Name = newSheetName;
        Save();
    }

    /// <inheritdoc />
    public string[] GetSheetNames()
    {
        var names = new List<string>();
        foreach (Interop.Worksheet sheet in _workbook.Worksheets) names.Add(sheet.Name);
        return [.. names];
    }

    /// <inheritdoc />
    public void CopyToWorkBook(string sheetName, string newSheetName)
    {
        var source = Sheet(sheetName);
        source.Copy(After: _workbook.Worksheets[_workbook.Worksheets.Count]);
        if (!string.IsNullOrWhiteSpace(newSheetName))
            ((Interop.Worksheet)_workbook.Worksheets[_workbook.Worksheets.Count]).Name = newSheetName;
        Save();
    }

    /// <inheritdoc />
    public void CopyToFile(string sheetName, CopyToFileRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        if (string.IsNullOrWhiteSpace(request.NewFilePath))
            throw new ArgumentException("NewFilePath is required.", nameof(request));

        var exists = File.Exists(request.NewFilePath);
        if (!exists && !request.AutoFileCreation)
        {
            throw new FileNotFoundException(
                $"'{request.NewFilePath}' does not exist. Set AutoFileCreation to create it.",
                request.NewFilePath);
        }

        var destination = exists
            ? _application.Workbooks.Open(request.NewFilePath,
                Password: request.NewFilePassword ?? string.Empty,
                WriteResPassword: request.NewModifyFilePassword ?? string.Empty)
            : _application.Workbooks.Add();

        Sheet(sheetName).Copy(After: destination.Worksheets[destination.Worksheets.Count]);
        if (!string.IsNullOrWhiteSpace(request.NewSheetName))
        {
            ((Interop.Worksheet)destination.Worksheets[destination.Worksheets.Count]).Name =
                request.NewSheetName;
        }

        if (exists) destination.Save();
        else destination.SaveAs(request.NewFilePath,
            Password: request.NewFilePassword ?? string.Empty,
            WriteResPassword: request.NewModifyFilePassword ?? string.Empty);
        destination.Close(SaveChanges: false);
    }

    /// <inheritdoc />
    public void CreateWorkBook(string sheetName)
    {
        var created = _application.Workbooks.Add();
        if (!string.IsNullOrWhiteSpace(sheetName))
            ((Interop.Worksheet)created.Worksheets[1]).Name = sheetName;
        created.SaveAs(FilePath);
        created.Close(SaveChanges: false);
    }

    /// <inheritdoc />
    public void SetSheetVisibility(string sheetName, XlSheetVisibility visibility)
    {
        Sheet(sheetName).Visible = visibility == XlSheetVisibility.Visible
            ? Interop.XlSheetVisibility.xlSheetVisible
            : Interop.XlSheetVisibility.xlSheetHidden;
        Save();
    }

    /// <inheritdoc />
    public void ProtectSheet(string sheetName, string password, ProtectUnProtectEnum type)
    {
        var sheet = Sheet(sheetName);
        if (type == ProtectUnProtectEnum.Protect) sheet.Protect(password ?? string.Empty);
        else sheet.Unprotect(password ?? string.Empty);
        Save();
    }

    /// <inheritdoc />
    public void SetPassword(string newPassword, string newModifyPassword)
    {
        _workbook.SaveAs(FilePath,
            Password: newPassword ?? string.Empty,
            WriteResPassword: newModifyPassword ?? string.Empty);
    }

    /// <inheritdoc />
    public void ExportWorkBook(string sheetName, string cellRange, string exportPath,
                               FixedFormatTypeEnum formatType)
    {
        if (string.IsNullOrWhiteSpace(exportPath))
            throw new ArgumentException("ExportPath is required.", nameof(exportPath));

        var format = formatType == FixedFormatTypeEnum.PDF
            ? Interop.XlFixedFormatType.xlTypePDF
            : Interop.XlFixedFormatType.xlTypeXPS;

        if (string.IsNullOrWhiteSpace(cellRange))
        {
            _workbook.ExportAsFixedFormat(format, exportPath);
            return;
        }

        // A range export goes through the sheet's print area, which is the only way
        // Excel will narrow a fixed-format export to part of a sheet.
        var sheet = Sheet(sheetName);
        var previous = sheet.PageSetup.PrintArea;
        try
        {
            sheet.PageSetup.PrintArea = cellRange;
            sheet.ExportAsFixedFormat(format, exportPath);
        }
        finally
        {
            sheet.PageSetup.PrintArea = previous;
        }
    }

    // -------------------------------------------------------------- plumbing

    /// <summary>Reads a used range into a table.</summary>
    internal static DataTable ToDataTable(Interop.Range range, bool hasHeader)
    {
        var table = new DataTable("Clipboard");
        var rows = range.Rows.Count;
        var columns = range.Columns.Count;

        for (var c = 1; c <= columns; c++)
        {
            var name = hasHeader
                ? ((Interop.Range)range.Cells[1, c]).Value2?.ToString()
                : null;
            if (string.IsNullOrWhiteSpace(name)) name = $"Column{c}";
            while (table.Columns.Contains(name)) name += "_";
            table.Columns.Add(name, typeof(string));
        }

        for (var r = hasHeader ? 2 : 1; r <= rows; r++)
        {
            var values = new object[columns];
            for (var c = 1; c <= columns; c++)
                values[c - 1] = ((Interop.Range)range.Cells[r, c]).Value2?.ToString() ?? string.Empty;
            table.Rows.Add(values);
        }
        return table;
    }

    private Interop.Worksheet Sheet(string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName)) return (Interop.Worksheet)_workbook.ActiveSheet;

        foreach (Interop.Worksheet sheet in _workbook.Worksheets)
        {
            if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase)) return sheet;
        }
        throw new ArgumentException(
            $"The workbook has no sheet named '{sheetName}'.", nameof(sheetName));
    }

    private Interop.Range Cell(string sheetName, string cell)
    {
        if (string.IsNullOrWhiteSpace(cell))
            throw new ArgumentException("Cell is required.", nameof(cell));
        return Sheet(sheetName).Range[cell];
    }

    private static void RequireFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("ImagePath is required.", nameof(path));
        if (!File.Exists(path))
            throw new FileNotFoundException($"The image '{path}' does not exist.", path);
    }

    private void Save()
    {
        if (File.Exists(FilePath)) _workbook.Save();
        else _workbook.SaveAs(FilePath);
    }

    private void Quit()
    {
        try { _application.Quit(); } catch (COMException) { /* already gone */ }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_closed) return;
        _closed = true;
        try { _workbook?.Close(SaveChanges: false); }
        catch (COMException) { /* nothing to close */ }
        Quit();
    }
}
