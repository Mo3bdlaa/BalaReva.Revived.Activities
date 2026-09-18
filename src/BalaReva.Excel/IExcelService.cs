using System.Data;
using BalaReva.Excel.Base;
using BalaReva.Excel.Charts;
using BalaReva.Excel.Enums;
using BalaReva.Excel.Sheets;
using BalaReva.Excel.Utilities;

namespace BalaReva.Excel;

/// <summary>
/// Opening workbooks, and the one operation that needs no workbook.
/// </summary>
/// <remarks>
/// The activities talk to this rather than to Excel's COM object model, so argument
/// handling and output mapping can be tested with a stand-in. No build agent has Excel
/// installed, so <see cref="ExcelService"/> itself is not covered by any automated test.
/// </remarks>
public interface IExcelService
{
    /// <summary>Opens a workbook, creating it when the path does not exist.</summary>
    IExcelWorkbook Open(string filePath, string filePassword, string modifyPassword);

    /// <summary>Reads whatever tabular data is on the clipboard.</summary>
    DataTable ClipboardToDataTable(bool hasHeader);
}

/// <summary>An open workbook and everything the activities do to it.</summary>
public interface IExcelWorkbook : IDisposable
{
    // charts

    /// <summary>Draws a chart on a sheet.</summary>
    void DrawChart(string sheetName, ChartRequest request);

    // comments

    /// <summary>Attaches a comment to a cell, replacing any already there.</summary>
    void AddComment(string sheetName, string cell, string comment);

    /// <summary>Removes a cell's comment.</summary>
    void DeleteComment(string sheetName, string cell);

    /// <summary>Reads a cell's comment, or empty when it has none.</summary>
    string GetComment(string sheetName, string cell);

    /// <summary>Shows or hides a cell's comment.</summary>
    void ShowHideComment(string sheetName, string cell, bool show);

    // data

    /// <summary>Copies a range to the clipboard.</summary>
    void CopyData(string sheetName, string copyRange);

    /// <summary>Clears a range.</summary>
    void DeleteData(string sheetName, string deleteRange);

    /// <summary>Finds and replaces text within a range.</summary>
    void FindReplace(string sheetName, string cellRange, string find, string replace,
                     FindReplaceEnum option);

    /// <summary>Removes duplicate rows from a range.</summary>
    void RemoveDuplicates(string sheetName, string cellRange, object[] columns, bool hasHeader);

    // layout

    /// <summary>Hides or shows named columns.</summary>
    void HideColumns(string sheetName, string[] columnNames, bool hide);

    /// <summary>Hides or shows numbered rows.</summary>
    void HideRows(string sheetName, int[] rowNumbers, bool hide);

    /// <summary>Autofits or sets the width of columns.</summary>
    void AutoFitColumns(string sheetName, string[] columnsRange, bool autoFit, double columnWidth);

    /// <summary>Autofits or sets the height of rows.</summary>
    void AutoFitRows(string sheetName, int[] rowsRange, bool autoFit, double rowHeight);

    /// <summary>Merges a range into one cell.</summary>
    void MergeCells(string sheetName, string mergeRange, string cellText,
                    AlignmentEnum horizontal, AlignmentEnum vertical);

    /// <summary>Splits a merged range back into cells.</summary>
    void UnMergeCells(string sheetName, string unMergeRange,
                      AlignmentEnum horizontal, AlignmentEnum vertical);

    /// <summary>Applies alignment and text control to ranges.</summary>
    void FormatCells(string sheetName, string[] cellRanges, CellFormatRequest request);

    /// <summary>Sets a cell's number format.</summary>
    void ChangeCellType(string sheetName, string cell, string cellFormat);

    /// <summary>Applies a built-in or named table style to a range.</summary>
    void InsertTableFormat(string sheetName, string cellRange,
                           TableFormatEnum style, string customStyle);

    /// <summary>Applies a built-in or named table style to an existing table.</summary>
    void SetTableFormat(string sheetName, int tableIndex,
                        TableFormatEnum style, string customStyle);

    // images and links

    /// <summary>Places an image on a sheet at a given position and size.</summary>
    void InsertImage(string sheetName, string imagePath, IObjectSize size);

    /// <summary>Places an image anchored to a cell.</summary>
    void InsertImageAtCell(string sheetName, string cell, string imagePath,
                           float imageWidth, float imageHeight);

    /// <summary>Saves each chart on a sheet as an image file.</summary>
    void ExtractGraphImage(string sheetName, string imageFolder, FileExtensEnum extension);

    /// <summary>Turns column values into hyperlinks.</summary>
    void AddHyperlinks(string sheetName, DataTable inputTable, string[] columnNames, bool overwriteText);

    /// <summary>Removes the hyperlinks from a range.</summary>
    void RemoveHyperlink(string sheetName, string cellRange);

    // workbook and sheets

    /// <summary>Adds a sheet with the given name.</summary>
    void AddSheet(string sheetName);

    /// <summary>Deletes a sheet.</summary>
    void DeleteSheet(string sheetName);

    /// <summary>Renames a sheet.</summary>
    void RenameSheet(string sheetName, string newSheetName);

    /// <summary>Names of every sheet in the workbook.</summary>
    string[] GetSheetNames();

    /// <summary>Copies a sheet within the same workbook.</summary>
    void CopyToWorkBook(string sheetName, string newSheetName);

    /// <summary>Copies a sheet into another workbook.</summary>
    void CopyToFile(string sheetName, CopyToFileRequest request);

    /// <summary>Creates a new workbook holding one named sheet.</summary>
    void CreateWorkBook(string sheetName);

    /// <summary>Shows or hides a sheet.</summary>
    void SetSheetVisibility(string sheetName, XlSheetVisibility visibility);

    /// <summary>Protects or unprotects a sheet.</summary>
    void ProtectSheet(string sheetName, string password, ProtectUnProtectEnum type);

    /// <summary>Sets the workbook's open and modify passwords.</summary>
    void SetPassword(string newPassword, string newModifyPassword);

    /// <summary>Exports the workbook, or a range of it, to PDF or XPS.</summary>
    void ExportWorkBook(string sheetName, string cellRange, string exportPath,
                        FixedFormatTypeEnum formatType);
}

/// <summary>Everything a chart activity passes to the service.</summary>
/// <remarks>Not part of the published surface; it keeps the service signature readable.</remarks>
public sealed class ChartRequest
{
    /// <summary>Excel's own chart type number, from one of the chart enums.</summary>
    public int ChartType { get; set; }

    /// <summary>Range holding the values to plot.</summary>
    public string CellRange { get; set; } = string.Empty;

    /// <summary>Range holding the category labels.</summary>
    public string LegendRange { get; set; } = string.Empty;

    /// <summary>Title shown above the chart.</summary>
    public string ChartTitle { get; set; } = string.Empty;

    /// <summary>Path to save a picture of the chart to. Empty saves none.</summary>
    public string ImageCopy { get; set; } = string.Empty;

    /// <summary>What the chart shows alongside its data.</summary>
    public ShowOptions Options { get; set; } = new();

    /// <summary>Where and how big the chart is.</summary>
    public ChartSize Size { get; set; } = new();

    /// <summary>Show the legend key beside each label. Column charts only.</summary>
    public bool ShowLegendKey { get; set; }

    /// <summary>Where value labels sit. Column charts only.</summary>
    public DataLabelPositionEnum ShowValuePosition { get; set; } = DataLabelPositionEnum.Center;

    /// <summary>Orientation of value labels. Column charts only.</summary>
    public TextOrientationEnum ShowValueTextOrientation { get; set; } = TextOrientationEnum.Horizontal;
}

/// <summary>Alignment and text control for <c>FormatCells</c>.</summary>
/// <remarks>Not part of the published surface; it keeps the service signature readable.</remarks>
public sealed class CellFormatRequest
{
    /// <summary>Horizontal alignment.</summary>
    public FormatHorizontalEnum Horizontal { get; set; } = FormatHorizontalEnum.General;

    /// <summary>Vertical alignment.</summary>
    public FormatVerticleEnum Verticle { get; set; } = FormatVerticleEnum.Bottom;

    /// <summary>Wrap text within the cell.</summary>
    public FormatTextControlEnum WrapText { get; set; } = FormatTextControlEnum.Select;

    /// <summary>Shrink text to fit the cell.</summary>
    public FormatTextControlEnum ShrinkFit { get; set; } = FormatTextControlEnum.Select;

    /// <summary>Merge the range into one cell.</summary>
    public FormatTextControlEnum MergeCells { get; set; } = FormatTextControlEnum.Select;

    /// <summary>Reading order.</summary>
    public FormatTextDirection TextDirection { get; set; } = FormatTextDirection.Context;

    /// <summary>Preset text orientation.</summary>
    public TextOrientationEumn TextOrientation { get; set; } = TextOrientationEumn.Select;

    /// <summary>Rotation in degrees, when the orientation is an angle.</summary>
    public int OrientationDeg { get; set; }

    /// <summary>Indent level.</summary>
    public int Indent { get; set; }
}

/// <summary>Where <c>CopyToFile</c> writes a copied sheet.</summary>
/// <remarks>Not part of the published surface; it keeps the service signature readable.</remarks>
public sealed class CopyToFileRequest
{
    /// <summary>Workbook to copy the sheet into.</summary>
    public string NewFilePath { get; set; } = string.Empty;

    /// <summary>Name to give the copied sheet.</summary>
    public string NewSheetName { get; set; } = string.Empty;

    /// <summary>Password needed to open the destination workbook.</summary>
    public string NewFilePassword { get; set; } = string.Empty;

    /// <summary>Password needed to change the destination workbook.</summary>
    public string NewModifyFilePassword { get; set; } = string.Empty;

    /// <summary>Create the destination workbook when it does not exist.</summary>
    public bool AutoFileCreation { get; set; }
}
