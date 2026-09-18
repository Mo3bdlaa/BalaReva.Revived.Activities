using System.Activities;
using System.Activities.Statements;
using System.Data;
using System.Drawing;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Main;
using BalaReva.EasyExcel.Utilities;
using Interop = Microsoft.Office.Interop.Excel;

namespace BalaReva.EasyExcel.Tests;

/// <summary>A stand-in workbook that records what it was asked to do.</summary>
/// <remarks>
/// No CI agent has Excel installed, so this is the only way these activities get
/// exercised. It covers argument handling, which call is made with which values, and how
/// results are mapped back onto output arguments.
/// </remarks>
public sealed class FakeExcelService : IExcelService, IExcelWorkbook
{
    public List<string> Calls { get; } = [];

    public bool Disposed { get; private set; }

    public ExcelOpenRequest? LastOpen { get; private set; }

    public SaveSheetRequest? LastSaveSheet { get; private set; }

    public ChartRef? LastChart { get; private set; }

    public ImageRef? LastImage { get; private set; }

    public TableRef? LastTable { get; private set; }

    public ChartBounds? LastBounds { get; private set; }

    public CellFontRequest? LastFont { get; private set; }

    public ClearOptions? LastClear { get; private set; }

    public PageSetupRequest? LastPageSetup { get; private set; }

    public MergeRequest? LastMerge { get; private set; }

    public string LastSheetName { get; private set; } = string.Empty;

    public Exception? Throw { get; set; }

    private T Record<T>(string call, T result)
    {
        Calls.Add(call);
        if (Throw is not null) throw Throw;
        return result;
    }

    private void Note(string sheetName) => LastSheetName = sheetName;

    // IExcelService

    public IExcelWorkbook Open(ExcelOpenRequest request)
    {
        LastOpen = request;
        FilePath = request.FilePath;
        // Deliberately ignores Throw: that flag stands for an operation failing, and
        // failing the open would mean ContinueOnError never got to do its job.
        Calls.Add($"Open({request.FilePath})");
        return this;
    }

    public void CloseAllExcel() => Record("CloseAllExcel()", 0);

    public void SaveAsSheet(SaveSheetRequest request)
    {
        LastSaveSheet = request;
        Record($"SaveAsSheet({request.FileName}->{request.NewFileName},sheet={request.Sheet})", 0);
    }

    // IExcelWorkbook

    public string FilePath { get; private set; } = string.Empty;

    /// <summary>Always null: there is no COM workbook behind a stand-in.</summary>
    public object? ComWorkbook => null;

    public bool ExistsAddIns(string addInsName) =>
        Record($"ExistsAddIns({addInsName})", addInsName == "Solver");

    public Dictionary<string, string> GetAllAddins() =>
        Record("GetAllAddins()", new Dictionary<string, string>
        {
            ["Solver"] = @"C:\addins\solver.xlam",
            ["Analysis"] = @"C:\addins\analys32.xll",
        });

    public void ChartCopyToClipboard(ChartRef chart) => TrackChart("ChartCopyToClipboard", chart);

    public void ChartDelete(ChartRef chart) => TrackChart("ChartDelete", chart);

    public void ChartDeleteAll(string sheetName)
    {
        Note(sheetName);
        Record($"ChartDeleteAll({sheetName})", 0);
    }

    public void DeleteAllCharts(string sheetName)
    {
        Note(sheetName);
        Record($"DeleteAllCharts({sheetName})", 0);
    }

    public void ChartEmbedToPowerPoint(ChartRef chart, string pptFile, int slideIndex, ChartBounds bounds)
    {
        LastBounds = bounds;
        TrackChart($"ChartEmbedToPowerPoint({pptFile},slide={slideIndex})", chart);
    }

    public void ChartFormat(ChartRef chart, ChartBounds bounds)
    {
        LastBounds = bounds;
        TrackChart("ChartFormat", chart);
    }

    public void ChartImageExtract(string sheetName, string imageFolder, FileExtension extension)
    {
        Note(sheetName);
        Record($"ChartImageExtract({sheetName},{imageFolder},{extension})", 0);
    }

    public void ExtractGraphImage(string sheetName, string imageFolder, FileExtension extension)
    {
        Note(sheetName);
        Record($"ExtractGraphImage({sheetName},{imageFolder},{extension})", 0);
    }

    public long[] FindEmptyRows(string sheetName, long startRowIndex)
    {
        Note(sheetName);
        return Record($"FindEmptyRows({sheetName},{startRowIndex})", new long[] { 4, 9 });
    }

    public long[] DeleteEmptyRows(string sheetName, long startRowIndex)
    {
        Note(sheetName);
        return Record($"DeleteEmptyRows({sheetName},{startRowIndex})", new long[] { 4, 9 });
    }

    public long[] HideUnhideEmptyRows(string sheetName, long startRowIndex, HideDeleteEnum hideDelete)
    {
        Note(sheetName);
        return Record($"HideUnhideEmptyRows({sheetName},{startRowIndex},{hideDelete})", new long[] { 4 });
    }

    public double RangeFunction(string sheetName, string cellRange, RangeFunctionKind kind)
    {
        Note(sheetName);
        return Record($"RangeFunction({sheetName},{cellRange},{kind})", 42d);
    }

    public void FreezeColumns(string sheetName, int columns, FreezePanesEnum option)
    {
        Note(sheetName);
        Record($"FreezeColumns({sheetName},{columns},{option})", 0);
    }

    public void FreezeRows(string sheetName, int rows, FreezePanesEnum option)
    {
        Note(sheetName);
        Record($"FreezeRows({sheetName},{rows},{option})", 0);
    }

    public void FreezePanes(string sheetName, string cellRange, FreezePanesEnum option)
    {
        Note(sheetName);
        Record($"FreezePanes({sheetName},{cellRange},{option})", 0);
    }

    public string GetHyperlink(string sheetName, string cell)
    {
        Note(sheetName);
        return Record($"GetHyperlink({sheetName},{cell})", "https://example.invalid/");
    }

    public void InsertHyperlink(string sheetName, string cell, string address,
                                string displayText, bool overwriteDisplayText)
    {
        Note(sheetName);
        Record($"InsertHyperlink({sheetName},{cell},{address},{displayText},{overwriteDisplayText})", 0);
    }

    public void RemoveHyperlink(string sheetName, string cell)
    {
        Note(sheetName);
        Record($"RemoveHyperlink({sheetName},{cell})", 0);
    }

    public (string[] Addresses, DataTable Table) ExtractHyperLinks(string sheetName, string cellRange)
    {
        Note(sheetName);
        return Record($"ExtractHyperLinks({sheetName},{cellRange})",
                      (new[] { "https://a.invalid/", "https://b.invalid/" }, new DataTable("HyperLinks")));
    }

    public void CollapseAllGroup(string sheetName)
    {
        Note(sheetName);
        Record($"CollapseAllGroup({sheetName})", 0);
    }

    public void GroupRange(string sheetName, string cellRange, GroupEnum groupType)
    {
        Note(sheetName);
        Record($"GroupRange({sheetName},{cellRange},{groupType})", 0);
    }

    public void UnGroup(string sheetName, string cellRange, GroupEnum groupType)
    {
        Note(sheetName);
        Record($"UnGroup({sheetName},{cellRange},{groupType})", 0);
    }

    public void General(bool adaptiveMenus) => Record($"General({adaptiveMenus})", 0);

    public void AddTrustedLocation(string folderPath, string description) =>
        Record($"AddTrustedLocation({folderPath},{description})", 0);

    public (string[] Paths, DataTable Table) ListTrustedLocation() =>
        Record("ListTrustedLocation()",
               (new[] { @"C:\trusted" }, new DataTable("TrustedLocations")));

    public void CopyAsPicture(string sheetName, string cell, string imageFilePath)
    {
        Note(sheetName);
        Record($"CopyAsPicture({sheetName},{cell},{imageFilePath})", 0);
    }

    public bool ImageExists(ImageRef image)
    {
        LastImage = image;
        return Record($"ImageExists@{Describe(image)}", true);
    }

    public void ImageExtractor(string sheetName, string imageFolder, FileExtension extension)
    {
        Note(sheetName);
        Record($"ImageExtractor({sheetName},{imageFolder},{extension})", 0);
    }

    public bool ImageResize(ImageRef image, float width, float height)
    {
        LastImage = image;
        return Record($"ImageResize@{Describe(image)}({width}x{height})", true);
    }

    public int ImagesDelete(string sheetName, string[] imageNames, int[] imageIndexes)
    {
        Note(sheetName);
        return Record(
            $"ImagesDelete({sheetName},[{string.Join(",", imageNames)}],[{string.Join(",", imageIndexes)}])",
            imageNames.Length + imageIndexes.Length);
    }

    public int ImagesDeleteAll(string sheetName)
    {
        Note(sheetName);
        return Record($"ImagesDeleteAll({sheetName})", 3);
    }

    public void CellFont(string sheetName, string cellRange, CellFontRequest request)
    {
        Note(sheetName);
        LastFont = request;
        Record($"CellFont({sheetName},{cellRange})", 0);
    }

    public CellFontRequest GetRangeStyle(string sheetName, string cellRange)
    {
        Note(sheetName);
        return Record($"GetRangeStyle({sheetName},{cellRange})", new CellFontRequest
        {
            FontName = "Calibri",
            FontSize = 11,
            FontColor = Color.FromArgb(255, 0, 0),
            BackgroundColor = Color.FromArgb(0, 0, 255),
        });
    }

    public void ClearSheet(string sheetName, string cellRange, ClearOptions options)
    {
        Note(sheetName);
        LastClear = options;
        Record($"ClearSheet({sheetName},{cellRange})", 0);
    }

    public void ColumnMove(string sheetName, string sourceColumns, string destinationColumns)
    {
        Note(sheetName);
        Record($"ColumnMove({sheetName},{sourceColumns}->{destinationColumns})", 0);
    }

    public void CopyToClipboard(string sheetName, string cell, bool readFilter)
    {
        Note(sheetName);
        Record($"CopyToClipboard({sheetName},{cell},filter={readFilter})", 0);
    }

    public void PasteClipboard(string sheetName, string cell)
    {
        Note(sheetName);
        Record($"PasteClipboard({sheetName},{cell})", 0);
    }

    public void DeleteColumns(string sheetName, string[] columnsRange)
    {
        Note(sheetName);
        Record($"DeleteColumns({sheetName},[{string.Join(",", columnsRange)}])", 0);
    }

    public void DeleteRows(string sheetName, string[] rowRange)
    {
        Note(sheetName);
        Record($"DeleteRows({sheetName},[{string.Join(",", rowRange)}])", 0);
    }

    public void FillColor(string sheetName, string cellRange, Color background,
                          Color patternColor, PatternEnum pattern)
    {
        Note(sheetName);
        Record($"FillColor({sheetName},{cellRange},{background.R}/{background.G}/{background.B},{pattern})", 0);
    }

    public (string[] Matches, int[] RowIndexes) Find(string sheetName, string cellRange,
                                                     string findText, FindReplaceEnum option, bool matchCase)
    {
        Note(sheetName);
        return Record($"Find({sheetName},{cellRange},{findText},{option},{matchCase})",
                      (new[] { "Total", "Total" }, new[] { 3, 8 }));
    }

    public bool FindReplace(string sheetName, string cellRange, string findText, string replaceText,
                            FindReplaceEnum option, bool matchCase)
    {
        Note(sheetName);
        return Record(
            $"FindReplace({sheetName},{cellRange},{findText}->{replaceText},{option},{matchCase})", true);
    }

    public (int Index, string Name) FindLastColumn(string sheetName)
    {
        Note(sheetName);
        return Record($"FindLastColumn({sheetName})", (27, "AA"));
    }

    public int FindLastRow(string sheetName)
    {
        Note(sheetName);
        return Record($"FindLastRow({sheetName})", 120);
    }

    public bool FormatPainter(string sheetName, string cellRange,
                              string destinationSheet, string destinationRange)
    {
        Note(sheetName);
        return Record($"FormatPainter({sheetName},{cellRange}->{destinationSheet},{destinationRange})", true);
    }

    public bool IsMergedCell(string sheetName, string cell)
    {
        Note(sheetName);
        return Record($"IsMergedCell({sheetName},{cell})", true);
    }

    public void PageSetup(string sheetName, PageSetupRequest request)
    {
        Note(sheetName);
        LastPageSetup = request;
        Record($"PageSetup({sheetName})", 0);
    }

    public void RefreshAll() => Record("RefreshAll()", 0);

    public void RemoveFilter(string sheetName, int tableIndex)
    {
        Note(sheetName);
        Record($"RemoveFilter({sheetName},{tableIndex})", 0);
    }

    public void RemoveFilterNonTable(string sheetName)
    {
        Note(sheetName);
        Record($"RemoveFilterNonTable({sheetName})", 0);
    }

    public void SelectCell(string sheetName, string cellRange)
    {
        Note(sheetName);
        Record($"SelectCell({sheetName},{cellRange})", 0);
    }

    public void SetBorder(string sheetName, string cellRange, BorderEnum presets,
                          Interop.XlLineStyle lineStyle, double weight, Color color)
    {
        Note(sheetName);
        Record($"SetBorder({sheetName},{cellRange},{presets},{lineStyle},{weight})", 0);
    }

    public void Sorting(string sheetName, TableRef table, string[] columnNames,
                        int[] columnIndexes, SortingTableEnum order)
    {
        Note(sheetName);
        LastTable = table;
        Record($"Sorting@{Describe(table)}([{string.Join(",", columnNames)}],"
               + $"[{string.Join(",", columnIndexes)}],{order})", 0);
    }

    public void DeleteTable(string sheetName, TableRef table)
    {
        Note(sheetName);
        TrackTable("DeleteTable", table);
    }

    public (List<string> Names, DataTable Table) GetTableNames(string sheetName)
    {
        Note(sheetName);
        return Record($"GetTableNames({sheetName})",
                      (new List<string> { "Sales", "Costs" }, new DataTable("Tables")));
    }

    public void ResizeTable(string sheetName, TableRef table, string newRange)
    {
        Note(sheetName);
        TrackTable($"ResizeTable({newRange})", table);
    }

    public bool TableExists(string sheetName, string tableName)
    {
        Note(sheetName);
        return Record($"TableExists({sheetName},{tableName})", tableName == "Sales");
    }

    public void ExportToTextFile(string sheetName, string textFilePath)
    {
        Note(sheetName);
        Record($"ExportToTextFile({sheetName},{textFilePath})", 0);
    }

    public void MergeSheetByRow(string sheetName, MergeRequest request)
    {
        Note(sheetName);
        LastMerge = request;
        Record($"MergeSheetByRow({sheetName},{request.AppendFileName})", 0);
    }

    public void MoveSheet(string sheetName, string beforeSheet)
    {
        Note(sheetName);
        Record($"MoveSheet({sheetName},before={beforeSheet})", 0);
    }

    public void SaveAsWorkBook(string fileName, FileFormatEnum format, string password, string writeResPassword) =>
        Record($"SaveAsWorkBook({fileName},{format},{password},{writeResPassword})", 0);

    public string[] ShowVisibleSheet(VisibleInvisibleEnum visibleType) =>
        Record($"ShowVisibleSheet({visibleType})", new[] { "Sheet1", "Summary" });

    public void TabColor(string sheetName, Color color)
    {
        Note(sheetName);
        Record($"TabColor({sheetName},{color.R}/{color.G}/{color.B})", 0);
    }

    public void Dispose() => Disposed = true;

    private void TrackChart(string call, ChartRef chart)
    {
        LastChart = chart;
        LastSheetName = chart.SheetName;
        Record($"{call}@{Describe(chart)}", 0);
    }

    private void TrackTable(string call, TableRef table)
    {
        LastTable = table;
        Record($"{call}@{Describe(table)}", 0);
    }

    private static string Describe(ChartRef chart) =>
        string.IsNullOrWhiteSpace(chart.ChartName)
            ? $"{chart.SheetName}/chart{chart.ChartIndex}"
            : $"{chart.SheetName}/{chart.ChartName}";

    private static string Describe(ImageRef image) =>
        string.IsNullOrWhiteSpace(image.ImageName)
            ? $"{image.SheetName}/image{image.ImageIndex}"
            : $"{image.SheetName}/{image.ImageName}";

    private static string Describe(TableRef table) =>
        string.IsNullOrWhiteSpace(table.TableName) ? $"table{table.TableIndex}" : table.TableName;
}
