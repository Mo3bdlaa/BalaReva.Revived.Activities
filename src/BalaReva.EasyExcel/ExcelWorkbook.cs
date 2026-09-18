using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using BalaReva.EasyExcel.Utilities;
using Interop = Microsoft.Office.Interop.Excel;

namespace BalaReva.EasyExcel;

/// <summary>A workbook held open through Excel's COM automation.</summary>
public sealed partial class ExcelWorkbook : IExcelWorkbook
{
    private readonly Interop.Application _application;
    private readonly Interop.Workbook _workbook;
    private readonly bool _displayAlerts;
    private bool _closed;

    internal ExcelWorkbook(ExcelOpenRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.FilePath))
            throw new ArgumentException("FilePath is required.", nameof(request));

        FilePath = request.FilePath;
        _displayAlerts = request.DisplayAlerts;

        _application = new Interop.Application
        {
            DisplayAlerts = request.DisplayAlerts,
            Visible = request.Visible,
        };

        // AutomationSecurity is an MsoAutomationSecurity, which lives in office.dll, so it
        // is set late-bound with the number instead. See Mso in ExcelService.cs.
        ((dynamic)_application).AutomationSecurity =
            request.MacrosEnabled ? Mso.SecurityLow : Mso.SecurityForceDisable;

        try
        {
            _workbook = File.Exists(request.FilePath)
                ? _application.Workbooks.Open(
                    Filename: request.FilePath,
                    UpdateLinks: (int)request.UpdateAutoLinks,
                    ReadOnly: false,
                    Password: request.FilePassword,
                    WriteResPassword: request.ModifyPassword)
                : _application.Workbooks.Add();

            if (!File.Exists(request.FilePath)) _workbook.SaveAs(request.FilePath);
        }
        catch
        {
            Quit();
            throw;
        }
    }

    /// <inheritdoc />
    public string FilePath { get; }

    /// <inheritdoc />
    public object? ComWorkbook => _workbook;

    /// <summary>The sheet by name, or the active sheet when no name is given.</summary>
    private Interop.Worksheet Sheet(string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            return (Interop.Worksheet)_workbook.ActiveSheet;

        foreach (Interop.Worksheet sheet in _workbook.Worksheets)
            if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                return sheet;

        throw new ArgumentException(
            $"'{FilePath}' has no sheet called '{sheetName}'.", nameof(sheetName));
    }

    /// <summary>The named range, or the whole used range when no range is given.</summary>
    private static Interop.Range Target(Interop.Worksheet sheet, string cellRange) =>
        string.IsNullOrWhiteSpace(cellRange) ? sheet.UsedRange : sheet.Range[cellRange];

    private void Save()
    {
        if (File.Exists(FilePath)) _workbook.Save();
        else _workbook.SaveAs(FilePath);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_closed) return;
        _closed = true;

        try { _workbook.Close(SaveChanges: true); }
        catch (COMException) { /* nothing to close */ }

        Quit();
    }

    private void Quit()
    {
        try
        {
            _application.DisplayAlerts = _displayAlerts;
            _application.Quit();
        }
        catch (COMException) { /* already gone */ }
        finally
        {
            // Excel outlives its last RCW otherwise, and the process stays in memory with
            // the file locked.
            if (Marshal.IsComObject(_application)) Marshal.FinalReleaseComObject(_application);
        }
    }

    /// <summary>Writes one sheet out as its own workbook. Used by <c>SaveAsSheet</c>.</summary>
    internal void SaveSheetAs(string sheetName, string newFileName)
    {
        var sheet = Sheet(sheetName);
        sheet.Copy();

        var copy = _application.ActiveWorkbook;
        copy.SaveAs(newFileName);
        copy.Close(SaveChanges: false);
    }

    // ---------------------------------------------------------------- add-ins

    /// <inheritdoc />
    public bool ExistsAddIns(string addInsName)
    {
        foreach (Interop.AddIn addIn in _application.AddIns)
            if (string.Equals(addIn.Name, addInsName, StringComparison.OrdinalIgnoreCase))
                return true;
        return false;
    }

    /// <inheritdoc />
    public Dictionary<string, string> GetAllAddins()
    {
        var addins = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (Interop.AddIn addIn in _application.AddIns)
            addins[addIn.Name] = addIn.FullName;
        return addins;
    }

    // ----------------------------------------------------------------- charts

    /// <summary>The chart on a sheet, by name when given and by position otherwise.</summary>
    private Interop.ChartObject ChartOn(ChartRef chart)
    {
        ArgumentNullException.ThrowIfNull(chart);
        var objects = (Interop.ChartObjects)Sheet(chart.SheetName).ChartObjects();

        if (!string.IsNullOrWhiteSpace(chart.ChartName))
        {
            foreach (Interop.ChartObject candidate in objects)
                if (string.Equals(candidate.Name, chart.ChartName, StringComparison.OrdinalIgnoreCase))
                    return candidate;

            throw new ArgumentException(
                $"No chart called '{chart.ChartName}'.", nameof(chart));
        }

        var index = chart.ChartIndex < 1 ? 1 : chart.ChartIndex;
        if (index > objects.Count)
            throw new ArgumentOutOfRangeException(
                nameof(chart), index, $"The sheet has {objects.Count} chart(s).");

        return (Interop.ChartObject)objects.Item(index);
    }

    /// <inheritdoc />
    public void ChartCopyToClipboard(ChartRef chart) => ChartOn(chart).Copy();

    /// <inheritdoc />
    public void ChartDelete(ChartRef chart)
    {
        ChartOn(chart).Delete();
        Save();
    }

    /// <inheritdoc />
    public void ChartDeleteAll(string sheetName)
    {
        ((Interop.ChartObjects)Sheet(sheetName).ChartObjects()).Delete();
        Save();
    }

    /// <inheritdoc />
    public void DeleteAllCharts(string sheetName) => ChartDeleteAll(sheetName);

    /// <inheritdoc />
    public void ChartFormat(ChartRef chart, ChartBounds bounds)
    {
        ArgumentNullException.ThrowIfNull(bounds);
        var target = ChartOn(chart);

        // Zero means "leave this measurement as it is", so each is set only when given.
        if (bounds.Left > 0) target.Left = bounds.Left;
        if (bounds.Top > 0) target.Top = bounds.Top;
        if (bounds.Width > 0) target.Width = bounds.Width;
        if (bounds.Height > 0) target.Height = bounds.Height;
        Save();
    }

    /// <inheritdoc />
    public void ChartEmbedToPowerPoint(ChartRef chart, string pptFile, int slideIndex, ChartBounds bounds)
    {
        ArgumentNullException.ThrowIfNull(bounds);
        if (string.IsNullOrWhiteSpace(pptFile))
            throw new ArgumentException("PptFile is required.", nameof(pptFile));

        ChartOn(chart).Copy();

        // PowerPoint is driven late-bound here for the same reason the EasyPowerPoint
        // package is driven late-bound throughout: its interop assembly needs office.dll,
        // which Microsoft publishes nowhere. See docs/REVIVAL.md.
        var progId = Type.GetTypeFromProgID("PowerPoint.Application")
            ?? throw new InvalidOperationException(
                "PowerPoint is not installed, or its COM automation server is not registered.");

        dynamic powerPoint = Activator.CreateInstance(progId)
            ?? throw new InvalidOperationException("PowerPoint's automation server could not be started.");

        try
        {
            dynamic presentation = File.Exists(pptFile)
                ? powerPoint.Presentations.Open(pptFile, ReadOnly: 0, Untitled: 0, WithWindow: 0)
                : powerPoint.Presentations.Add(0);

            var index = slideIndex < 1 ? 1 : slideIndex;
            while (presentation.Slides.Count < index) presentation.Slides.Add(presentation.Slides.Count + 1, 12);

            dynamic shape = presentation.Slides[index].Shapes.Paste()[1];
            if (bounds.Left > 0) shape.Left = (float)bounds.Left;
            if (bounds.Top > 0) shape.Top = (float)bounds.Top;
            if (bounds.Width > 0) shape.Width = (float)bounds.Width;
            if (bounds.Height > 0) shape.Height = (float)bounds.Height;

            if (File.Exists(pptFile)) presentation.Save();
            else presentation.SaveAs(pptFile);
            presentation.Close();
        }
        finally
        {
            try { powerPoint.Quit(); } catch (COMException) { /* already gone */ }
        }
    }

    /// <inheritdoc />
    public void ChartImageExtract(string sheetName, string imageFolder, FileExtension extension) =>
        ExtractGraphImage(sheetName, imageFolder, extension);

    /// <inheritdoc />
    public void ExtractGraphImage(string sheetName, string imageFolder, FileExtension extension)
    {
        Directory.CreateDirectory(imageFolder);
        var sheet = Sheet(sheetName);
        var suffix = extension.ToString().ToLowerInvariant();
        var index = 0;

        foreach (Interop.ChartObject chart in (Interop.ChartObjects)sheet.ChartObjects())
        {
            index++;
            var name = string.IsNullOrWhiteSpace(chart.Name) ? $"Chart{index}" : chart.Name;
            chart.Chart.Export(Path.Combine(imageFolder, $"{name}.{suffix}"), suffix.ToUpperInvariant());
        }
    }
}
