using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using BalaReva.EasyExcel.Utilities;
using Interop = Microsoft.Office.Interop.Excel;

namespace BalaReva.EasyExcel;

public sealed partial class ExcelWorkbook
{
    // ------------------------------------------------------------ empty rows

    /// <inheritdoc />
    public long[] FindEmptyRows(string sheetName, long startRowIndex) =>
        [.. EmptyRows(Sheet(sheetName), startRowIndex)];

    /// <inheritdoc />
    public long[] DeleteEmptyRows(string sheetName, long startRowIndex)
    {
        var sheet = Sheet(sheetName);
        var rows = EmptyRows(sheet, startRowIndex);

        // Bottom up, so deleting one does not renumber the ones still to go.
        foreach (var row in rows.OrderByDescending(r => r))
            ((Interop.Range)sheet.Rows[row]).Delete();

        Save();
        return [.. rows];
    }

    /// <inheritdoc />
    public long[] HideUnhideEmptyRows(string sheetName, long startRowIndex, HideDeleteEnum hideDelete)
    {
        var sheet = Sheet(sheetName);
        var rows = EmptyRows(sheet, startRowIndex);

        foreach (var row in rows)
            ((Interop.Range)sheet.Rows[row]).EntireRow.Hidden = hideDelete == HideDeleteEnum.Hide;

        Save();
        return [.. rows];
    }

    private static List<long> EmptyRows(Interop.Worksheet sheet, long startRowIndex)
    {
        var used = sheet.UsedRange;
        var first = Math.Max(startRowIndex, 1);
        var last = used.Row + used.Rows.Count - 1;
        var empty = new List<long>();

        for (var row = first; row <= last; row++)
        {
            var range = (Interop.Range)sheet.Rows[row];
            if (sheet.Application.WorksheetFunction.CountA(range) == 0) empty.Add(row);
        }
        return empty;
    }

    // -------------------------------------------------------------- formulas

    /// <inheritdoc />
    public double RangeFunction(string sheetName, string cellRange, RangeFunctionKind kind)
    {
        var range = Target(Sheet(sheetName), cellRange);
        var functions = _application.WorksheetFunction;

        return kind switch
        {
            RangeFunctionKind.Average => functions.Average(range),
            RangeFunctionKind.Count => functions.Count(range),
            RangeFunctionKind.CountA => functions.CountA(range),
            RangeFunctionKind.Max => functions.Max(range),
            RangeFunctionKind.Min => functions.Min(range),
            RangeFunctionKind.Sum => functions.Sum(range),
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unknown range function."),
        };
    }

    // ---------------------------------------------------------- freeze panes

    /// <inheritdoc />
    public void FreezeColumns(string sheetName, int columns, FreezePanesEnum option) =>
        Freeze(sheetName, $"{ColumnName(Math.Max(columns, 1) + 1)}1", option);

    /// <inheritdoc />
    public void FreezeRows(string sheetName, int rows, FreezePanesEnum option) =>
        Freeze(sheetName, $"A{Math.Max(rows, 1) + 1}", option);

    /// <inheritdoc />
    public void FreezePanes(string sheetName, string cellRange, FreezePanesEnum option) =>
        Freeze(sheetName, cellRange, option);

    /// <summary>
    /// Freezing works on the active window, so the sheet has to be activated first.
    /// </summary>
    private void Freeze(string sheetName, string cell, FreezePanesEnum option)
    {
        var sheet = Sheet(sheetName);
        sheet.Activate();

        var window = _application.ActiveWindow;
        window.FreezePanes = false;

        if (option == FreezePanesEnum.Freeze)
        {
            sheet.Range[cell].Select();
            window.FreezePanes = true;
        }
        Save();
    }

    /// <summary>Turns a column number into its letters. Excel counts from 1, with no zero digit.</summary>
    private static string ColumnName(int index)
    {
        var name = string.Empty;
        while (index > 0)
        {
            var digit = (index - 1) % 26;
            name = (char)('A' + digit) + name;
            index = (index - 1) / 26;
        }
        return name;
    }

    // ------------------------------------------------------------ hyperlinks

    /// <inheritdoc />
    public string GetHyperlink(string sheetName, string cell)
    {
        var links = Sheet(sheetName).Range[cell].Hyperlinks;
        return links.Count == 0 ? string.Empty : links[1].Address ?? string.Empty;
    }

    /// <inheritdoc />
    public void InsertHyperlink(string sheetName, string cell, string address,
                                string displayText, bool overwriteDisplayText)
    {
        var sheet = Sheet(sheetName);
        var range = sheet.Range[cell];
        var text = overwriteDisplayText || string.IsNullOrEmpty(Convert.ToString(range.Text))
            ? (string.IsNullOrWhiteSpace(displayText) ? address : displayText)
            : Convert.ToString(range.Text)!;

        sheet.Hyperlinks.Add(range, address, Type.Missing, Type.Missing, text);
        Save();
    }

    /// <inheritdoc />
    public void RemoveHyperlink(string sheetName, string cell)
    {
        Sheet(sheetName).Range[cell].Hyperlinks.Delete();
        Save();
    }

    /// <inheritdoc />
    public (string[] Addresses, DataTable Table) ExtractHyperLinks(string sheetName, string cellRange)
    {
        var table = new DataTable("HyperLinks");
        table.Columns.Add("Cell", typeof(string));
        table.Columns.Add("TextToDisplay", typeof(string));
        table.Columns.Add("Address", typeof(string));

        var addresses = new List<string>();
        foreach (Interop.Hyperlink link in Target(Sheet(sheetName), cellRange).Hyperlinks)
        {
            var address = link.Address ?? string.Empty;
            addresses.Add(address);
            table.Rows.Add(link.Range.Address[false, false], link.TextToDisplay ?? string.Empty, address);
        }
        return ([.. addresses], table);
    }

    // ----------------------------------------------------------------- outline

    /// <inheritdoc />
    public void CollapseAllGroup(string sheetName)
    {
        Sheet(sheetName).Outline.ShowLevels(RowLevels: 1, ColumnLevels: 1);
        Save();
    }

    /// <inheritdoc />
    public void GroupRange(string sheetName, string cellRange, GroupEnum groupType)
    {
        Grouped(sheetName, cellRange, groupType).Group();
        Save();
    }

    /// <inheritdoc />
    public void UnGroup(string sheetName, string cellRange, GroupEnum groupType)
    {
        Grouped(sheetName, cellRange, groupType).Ungroup();
        Save();
    }

    private Interop.Range Grouped(string sheetName, string cellRange, GroupEnum groupType)
    {
        var range = Sheet(sheetName).Range[cellRange];
        return groupType == GroupEnum.Rows ? range.EntireRow : range.EntireColumn;
    }

    // ---------------------------------------------------------------- settings

    /// <inheritdoc />
    public void General(bool adaptiveMenus) =>
        // CommandBars is an Office type, so this goes through dynamic like the rest of
        // the office.dll surface. See Mso in ExcelService.cs.
        ((dynamic)_application).CommandBars.AdaptiveMenus = adaptiveMenus;

    /// <inheritdoc />
    public void AddTrustedLocation(string folderPath, string description)
    {
        if (!Directory.Exists(folderPath))
            throw new DirectoryNotFoundException($"'{folderPath}' does not exist.");

        // Application.AddTrustedLocation is not on the typed interop surface, so it goes
        // through dynamic. Excel refuses it outright under some Trust Center policies,
        // which is worth saying plainly rather than passing a COM error up.
        try
        {
            ((dynamic)_application).AddTrustedLocation(folderPath, description, true);
        }
        catch (Exception error) when (error is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException
                                               or COMException)
        {
            throw new NotSupportedException(
                "Excel would not add a trusted location through automation. Trust Center "
                + "policy can forbid this. Add the folder under File, Options, Trust Center.",
                error);
        }
    }

    /// <inheritdoc />
    public (string[] Paths, DataTable Table) ListTrustedLocation()
    {
        var table = new DataTable("TrustedLocations");
        table.Columns.Add("Path", typeof(string));
        table.Columns.Add("Description", typeof(string));
        table.Columns.Add("AllowSubFolders", typeof(bool));

        var paths = new List<string>();
        foreach (dynamic location in ((dynamic)_application).TrustedLocations)
        {
            string path = location.Path;
            paths.Add(path);
            table.Rows.Add(path, (string)location.Description, (bool)location.AllowSubfolders);
        }
        return ([.. paths], table);
    }
}
