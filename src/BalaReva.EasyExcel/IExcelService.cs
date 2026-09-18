using System.Data;
using System.Drawing;
using BalaReva.EasyExcel.Utilities;
using Interop = Microsoft.Office.Interop.Excel;

namespace BalaReva.EasyExcel;

/// <summary>Opens workbooks, and does the few things that do not need one.</summary>
/// <remarks>
/// The activities talk to this rather than to Excel's COM object model, so argument
/// handling and output mapping can be tested with a stand-in. No build agent has Excel
/// installed, so <see cref="ExcelService"/> itself is not covered by any automated test.
/// </remarks>
public interface IExcelService
{
    /// <summary>Opens a workbook.</summary>
    IExcelWorkbook Open(ExcelOpenRequest request);

    /// <summary>Closes every running Excel instance, saving nothing.</summary>
    void CloseAllExcel();

    /// <summary>Saves one sheet of a workbook out as its own file.</summary>
    void SaveAsSheet(SaveSheetRequest request);
}

/// <summary>Reads a workbook through the Open XML package format, without Excel.</summary>
public interface IOpenXmlReader
{
    /// <summary>Column letters hidden on a sheet.</summary>
    List<string> HiddenColumns(string filePath, string sheetName);

    /// <summary>Row numbers hidden on a sheet, numbered from 1.</summary>
    List<int> HiddenRows(string filePath, string sheetName);
}

/// <summary>An open workbook and everything the activities do to it.</summary>
public interface IExcelWorkbook : IDisposable
{
    /// <summary>Full path of the open workbook.</summary>
    string FilePath { get; }

    /// <summary>
    /// The underlying COM workbook, or null when there is not one.
    /// </summary>
    /// <remarks>
    /// Typed as object so a stand-in need not produce one; the scope casts it to hand a
    /// workflow the interop type the published package exposed on
    /// <c>ExcelParam.ExcelWorkBook</c>.
    /// </remarks>
    object? ComWorkbook { get; }

    // add-ins

    /// <summary>Whether an add-in of that name is installed.</summary>
    bool ExistsAddIns(string addInsName);

    /// <summary>Every add-in, by name and full path.</summary>
    Dictionary<string, string> GetAllAddins();

    // charts

    /// <summary>Copies a chart to the clipboard.</summary>
    void ChartCopyToClipboard(ChartRef chart);

    /// <summary>Deletes a chart.</summary>
    void ChartDelete(ChartRef chart);

    /// <summary>Deletes every chart on a sheet.</summary>
    void ChartDeleteAll(string sheetName);

    /// <summary>Copies a chart onto a slide of a presentation.</summary>
    void ChartEmbedToPowerPoint(ChartRef chart, string pptFile, int slideIndex, ChartBounds bounds);

    /// <summary>Moves and resizes a chart, in points. Zero leaves a measurement alone.</summary>
    void ChartFormat(ChartRef chart, ChartBounds bounds);

    /// <summary>Saves a sheet's charts as images.</summary>
    void ChartImageExtract(string sheetName, string imageFolder, FileExtension extension);

    /// <summary>Deletes every chart on a sheet.</summary>
    void DeleteAllCharts(string sheetName);

    // empty rows

    /// <summary>Finds the empty rows from a starting row onwards.</summary>
    long[] FindEmptyRows(string sheetName, long startRowIndex);

    /// <summary>Deletes the empty rows from a starting row onwards, and reports them.</summary>
    long[] DeleteEmptyRows(string sheetName, long startRowIndex);

    /// <summary>Hides or deletes the empty rows from a starting row onwards.</summary>
    long[] HideUnhideEmptyRows(string sheetName, long startRowIndex, HideDeleteEnum hideDelete);

    // formulas

    /// <summary>Evaluates one of Excel's range functions over a range.</summary>
    double RangeFunction(string sheetName, string cellRange, RangeFunctionKind kind);

    // freeze panes

    /// <summary>Freezes or unfreezes the first columns of a sheet.</summary>
    void FreezeColumns(string sheetName, int columns, FreezePanesEnum option);

    /// <summary>Freezes or unfreezes the first rows of a sheet.</summary>
    void FreezeRows(string sheetName, int rows, FreezePanesEnum option);

    /// <summary>Freezes or unfreezes the panes at a cell.</summary>
    void FreezePanes(string sheetName, string cellRange, FreezePanesEnum option);

    // hyperlinks

    /// <summary>The address a cell's hyperlink points at, or empty when it has none.</summary>
    string GetHyperlink(string sheetName, string cell);

    /// <summary>Puts a hyperlink on a cell.</summary>
    void InsertHyperlink(string sheetName, string cell, string address,
                         string displayText, bool overwriteDisplayText);

    /// <summary>Removes a cell's hyperlink.</summary>
    void RemoveHyperlink(string sheetName, string cell);

    /// <summary>Every hyperlink in a range, as addresses and as a table.</summary>
    (string[] Addresses, DataTable Table) ExtractHyperLinks(string sheetName, string cellRange);

    // outline

    /// <summary>Collapses every outline group on a sheet.</summary>
    void CollapseAllGroup(string sheetName);

    /// <summary>Groups a range by row or by column.</summary>
    void GroupRange(string sheetName, string cellRange, GroupEnum groupType);

    /// <summary>Ungroups a range by row or by column.</summary>
    void UnGroup(string sheetName, string cellRange, GroupEnum groupType);

    // settings

    /// <summary>Turns Excel's adaptive menus on or off.</summary>
    void General(bool adaptiveMenus);

    /// <summary>Adds a folder to Excel's trusted locations.</summary>
    void AddTrustedLocation(string folderPath, string description);

    /// <summary>Every trusted location, as paths and as a table.</summary>
    (string[] Paths, DataTable Table) ListTrustedLocation();

    // sheet images

    /// <summary>Copies a range to an image file.</summary>
    void CopyAsPicture(string sheetName, string cell, string imageFilePath);

    /// <summary>Whether a picture of that name or at that position exists.</summary>
    bool ImageExists(ImageRef image);

    /// <summary>Saves a sheet's pictures as image files.</summary>
    void ImageExtractor(string sheetName, string imageFolder, FileExtension extension);

    /// <summary>Resizes a picture, in points. Zero leaves a measurement alone.</summary>
    bool ImageResize(ImageRef image, float width, float height);

    /// <summary>Deletes the named or numbered pictures, and reports how many went.</summary>
    int ImagesDelete(string sheetName, string[] imageNames, int[] imageIndexes);

    /// <summary>Deletes every picture on a sheet, and reports how many went.</summary>
    int ImagesDeleteAll(string sheetName);

    // sheets

    /// <summary>Applies font and fill to a range.</summary>
    void CellFont(string sheetName, string cellRange, CellFontRequest request);

    /// <summary>Reads the font and fill of a range's first cell.</summary>
    CellFontRequest GetRangeStyle(string sheetName, string cellRange);

    /// <summary>Clears the named parts of a range.</summary>
    void ClearSheet(string sheetName, string cellRange, ClearOptions options);

    /// <summary>Moves a column range somewhere else on the sheet.</summary>
    void ColumnMove(string sheetName, string sourceColumns, string destinationColumns);

    /// <summary>Copies a range to the clipboard.</summary>
    void CopyToClipboard(string sheetName, string cell, bool readFilter);

    /// <summary>Pastes the clipboard at a cell.</summary>
    void PasteClipboard(string sheetName, string cell);

    /// <summary>Deletes whole columns.</summary>
    void DeleteColumns(string sheetName, string[] columnsRange);

    /// <summary>Deletes whole rows.</summary>
    void DeleteRows(string sheetName, string[] rowRange);

    /// <summary>Saves a sheet's charts as images.</summary>
    void ExtractGraphImage(string sheetName, string imageFolder, FileExtension extension);

    /// <summary>Fills a range with a colour and a pattern.</summary>
    void FillColor(string sheetName, string cellRange, Color background,
                   Color patternColor, PatternEnum pattern);

    /// <summary>Finds text in a range, reporting the matches and the rows they are on.</summary>
    (string[] Matches, int[] RowIndexes) Find(string sheetName, string cellRange,
                                              string findText, FindReplaceEnum option, bool matchCase);

    /// <summary>Replaces text in a range.</summary>
    bool FindReplace(string sheetName, string cellRange, string findText, string replaceText,
                     FindReplaceEnum option, bool matchCase);

    /// <summary>The last column of a sheet with anything in it.</summary>
    (int Index, string Name) FindLastColumn(string sheetName);

    /// <summary>The last row of a sheet with anything in it.</summary>
    int FindLastRow(string sheetName);

    /// <summary>Copies the formatting of one range onto another.</summary>
    bool FormatPainter(string sheetName, string cellRange,
                       string destinationSheet, string destinationRange);

    /// <summary>Whether a cell is part of a merged cell.</summary>
    bool IsMergedCell(string sheetName, string cell);

    /// <summary>Sets a sheet's page setup for printing.</summary>
    void PageSetup(string sheetName, PageSetupRequest request);

    /// <summary>Refreshes every data connection in the workbook.</summary>
    void RefreshAll();

    /// <summary>Clears the filter on a table.</summary>
    void RemoveFilter(string sheetName, int tableIndex);

    /// <summary>Clears the filter on a sheet that has no table.</summary>
    void RemoveFilterNonTable(string sheetName);

    /// <summary>Selects a range.</summary>
    void SelectCell(string sheetName, string cellRange);

    /// <summary>Draws borders around or through a range.</summary>
    void SetBorder(string sheetName, string cellRange, BorderEnum presets,
                   Interop.XlLineStyle lineStyle, double weight, Color color);

    /// <summary>Sorts a table by the named or numbered columns.</summary>
    void Sorting(string sheetName, TableRef table, string[] columnNames,
                 int[] columnIndexes, SortingTableEnum order);

    // tables

    /// <summary>Deletes a table, leaving its cells behind.</summary>
    void DeleteTable(string sheetName, TableRef table);

    /// <summary>Every table on a sheet, as names and as a table.</summary>
    (List<string> Names, DataTable Table) GetTableNames(string sheetName);

    /// <summary>Resizes a table to a new range.</summary>
    void ResizeTable(string sheetName, TableRef table, string newRange);

    /// <summary>Whether a table of that name exists on the sheet.</summary>
    bool TableExists(string sheetName, string tableName);

    // workbook

    /// <summary>Writes a sheet out as a text file.</summary>
    void ExportToTextFile(string sheetName, string textFilePath);

    /// <summary>Appends the rows of another workbook's sheet to one of this workbook's.</summary>
    void MergeSheetByRow(string sheetName, MergeRequest request);

    /// <summary>Moves a sheet in front of another.</summary>
    void MoveSheet(string sheetName, string beforeSheet);

    /// <summary>Saves a copy under another name and format.</summary>
    void SaveAsWorkBook(string fileName, FileFormatEnum format, string password, string writeResPassword);

    /// <summary>The sheets that are visible, or the ones that are not.</summary>
    string[] ShowVisibleSheet(VisibleInvisibleEnum visibleType);

    /// <summary>Colours a sheet's tab.</summary>
    void TabColor(string sheetName, Color color);
}
