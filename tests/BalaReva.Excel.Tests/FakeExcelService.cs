using System.Activities;
using System.Data;
using BalaReva.Excel.Base;
using BalaReva.Excel.Charts;
using BalaReva.Excel.Enums;
using BalaReva.Excel.Sheets;
using BalaReva.Excel.Utilities;

namespace BalaReva.Excel.Tests;

/// <summary>
/// A stand-in workbook that records what it was asked to do.
/// </summary>
/// <remarks>
/// No CI agent has Excel installed, so this is the only way these activities get
/// exercised. It covers argument handling, which call is made with which values, and
/// how results are mapped back onto output arguments.
/// </remarks>
public sealed class FakeExcelService : IExcelService, IExcelWorkbook
{
    public List<string> Calls { get; } = [];

    public bool Disposed { get; private set; }

    public string Comment { get; set; } = "a note";

    public string[] Sheets { get; set; } = ["Sheet1", "Data"];

    public DataTable Clipboard { get; set; } = new("Clipboard");

    public ChartRequest? LastChart { get; private set; }

    public CellFormatRequest? LastFormat { get; private set; }

    public CopyToFileRequest? LastCopy { get; private set; }

    public Exception? Throw { get; set; }

    private T Record<T>(string call, T result)
    {
        Calls.Add(call);
        if (Throw is not null) throw Throw;
        return result;
    }

    // IExcelService

    public IExcelWorkbook Open(string filePath, string filePassword, string modifyPassword)
    {
        // Does not honour Throw: the flag stands for an operation failing, and failing
        // the open would mean ContinueOnError never got to do its job.
        Calls.Add($"Open({filePath})");
        return this;
    }

    public DataTable ClipboardToDataTable(bool hasHeader) =>
        Record($"ClipboardToDataTable(header={hasHeader})", Clipboard);

    // IExcelWorkbook

    public void DrawChart(string sheetName, ChartRequest request)
    {
        LastChart = request;
        Record($"DrawChart({sheetName},{request.ChartType},{request.CellRange})", 0);
    }

    public void AddComment(string sheetName, string cell, string comment) =>
        Record($"AddComment({sheetName},{cell},{comment})", 0);

    public void DeleteComment(string sheetName, string cell) =>
        Record($"DeleteComment({sheetName},{cell})", 0);

    public string GetComment(string sheetName, string cell) =>
        Record($"GetComment({sheetName},{cell})", Comment);

    public void ShowHideComment(string sheetName, string cell, bool show) =>
        Record($"ShowHideComment({sheetName},{cell},{show})", 0);

    public void CopyData(string sheetName, string copyRange) =>
        Record($"CopyData({sheetName},{copyRange})", 0);

    public void DeleteData(string sheetName, string deleteRange) =>
        Record($"DeleteData({sheetName},{deleteRange})", 0);

    public void FindReplace(string sheetName, string cellRange, string find, string replace,
                            FindReplaceEnum option) =>
        Record($"FindReplace({sheetName},{cellRange},{find}->{replace},{option})", 0);

    public void RemoveDuplicates(string sheetName, string cellRange, object[] columns, bool hasHeader) =>
        Record($"RemoveDuplicates({sheetName},{cellRange},[{string.Join(",", columns)}],{hasHeader})", 0);

    public void HideColumns(string sheetName, string[] columnNames, bool hide) =>
        Record($"HideColumns({sheetName},[{string.Join(",", columnNames)}],{hide})", 0);

    public void HideRows(string sheetName, int[] rowNumbers, bool hide) =>
        Record($"HideRows({sheetName},[{string.Join(",", rowNumbers)}],{hide})", 0);

    public void AutoFitColumns(string sheetName, string[] columnsRange, bool autoFit, double columnWidth) =>
        Record($"AutoFitColumns({sheetName},[{string.Join(",", columnsRange)}],{autoFit},{columnWidth})", 0);

    public void AutoFitRows(string sheetName, int[] rowsRange, bool autoFit, double rowHeight) =>
        Record($"AutoFitRows({sheetName},[{string.Join(",", rowsRange)}],{autoFit},{rowHeight})", 0);

    public void MergeCells(string sheetName, string mergeRange, string cellText,
                           AlignmentEnum horizontal, AlignmentEnum vertical) =>
        Record($"MergeCells({sheetName},{mergeRange},{cellText},{horizontal},{vertical})", 0);

    public void UnMergeCells(string sheetName, string unMergeRange,
                             AlignmentEnum horizontal, AlignmentEnum vertical) =>
        Record($"UnMergeCells({sheetName},{unMergeRange},{horizontal},{vertical})", 0);

    public void FormatCells(string sheetName, string[] cellRanges, CellFormatRequest request)
    {
        LastFormat = request;
        Record($"FormatCells({sheetName},[{string.Join(",", cellRanges)}])", 0);
    }

    public void ChangeCellType(string sheetName, string cell, string cellFormat) =>
        Record($"ChangeCellType({sheetName},{cell},{cellFormat})", 0);

    public void InsertTableFormat(string sheetName, string cellRange,
                                  TableFormatEnum style, string customStyle) =>
        Record($"InsertTableFormat({sheetName},{cellRange},{style},{customStyle})", 0);

    public void SetTableFormat(string sheetName, int tableIndex,
                               TableFormatEnum style, string customStyle) =>
        Record($"SetTableFormat({sheetName},{tableIndex},{style},{customStyle})", 0);

    public void InsertImage(string sheetName, string imagePath, IObjectSize size) =>
        Record($"InsertImage({sheetName},{imagePath},{size.Left}x{size.Top})", 0);

    public void InsertImageAtCell(string sheetName, string cell, string imagePath,
                                  float imageWidth, float imageHeight) =>
        Record($"InsertImageAtCell({sheetName},{cell},{imagePath},{imageWidth}x{imageHeight})", 0);

    public void ExtractGraphImage(string sheetName, string imageFolder, FileExtensEnum extension) =>
        Record($"ExtractGraphImage({sheetName},{imageFolder},{extension})", 0);

    public void AddHyperlinks(string sheetName, DataTable inputTable, string[] columnNames, bool overwriteText) =>
        Record($"AddHyperlinks({sheetName},rows={inputTable.Rows.Count},"
               + $"[{string.Join(",", columnNames)}],{overwriteText})", 0);

    public void RemoveHyperlink(string sheetName, string cellRange) =>
        Record($"RemoveHyperlink({sheetName},{cellRange})", 0);

    public void AddSheet(string sheetName) => Record($"AddSheet({sheetName})", 0);

    public void DeleteSheet(string sheetName) => Record($"DeleteSheet({sheetName})", 0);

    public void RenameSheet(string sheetName, string newSheetName) =>
        Record($"RenameSheet({sheetName},{newSheetName})", 0);

    public string[] GetSheetNames() => Record("GetSheetNames", Sheets);

    public void CopyToWorkBook(string sheetName, string newSheetName) =>
        Record($"CopyToWorkBook({sheetName},{newSheetName})", 0);

    public void CopyToFile(string sheetName, CopyToFileRequest request)
    {
        LastCopy = request;
        Record($"CopyToFile({sheetName},{request.NewFilePath})", 0);
    }

    public void CreateWorkBook(string sheetName) => Record($"CreateWorkBook({sheetName})", 0);

    public void SetSheetVisibility(string sheetName, XlSheetVisibility visibility) =>
        Record($"SetSheetVisibility({sheetName},{visibility})", 0);

    public void ProtectSheet(string sheetName, string password, ProtectUnProtectEnum type) =>
        Record($"ProtectSheet({sheetName},{password},{type})", 0);

    public void SetPassword(string newPassword, string newModifyPassword) =>
        Record($"SetPassword({newPassword},{newModifyPassword})", 0);

    public void ExportWorkBook(string sheetName, string cellRange, string exportPath,
                               FixedFormatTypeEnum formatType) =>
        Record($"ExportWorkBook({sheetName},{cellRange},{exportPath},{formatType})", 0);

    public void Dispose() => Disposed = true;
}

/// <summary>
/// Runs one activity as a workflow root with a stand-in workbook registered.
/// </summary>
/// <remarks>
/// This package has no scope activity, so an activity is the root and its output
/// arguments come straight back from Invoke.
/// </remarks>
internal static class Harness
{
    public static IDictionary<string, object> Run(ExcelCore activity, FakeExcelService service,
                                                  string filePath = @"C:\books\data.xlsx")
    {
        activity.ContinueOnError ??= new InArgument<bool>(false);
        activity.Delay ??= new InArgument<short>(0);
        activity.FilePath ??= new InArgument<string>(filePath);

        var invoker = new WorkflowInvoker(activity);
        invoker.Extensions.Add(service);
        return invoker.Invoke();
    }
}
