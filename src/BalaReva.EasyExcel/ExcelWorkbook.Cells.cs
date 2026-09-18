using System.Data;
using System.Drawing;
using BalaReva.EasyExcel.Utilities;
using Interop = Microsoft.Office.Interop.Excel;

namespace BalaReva.EasyExcel;

public sealed partial class ExcelWorkbook
{
    /// <summary>Excel wants a BGR integer where .NET gives RGB.</summary>
    private static int Bgr(Color color) => color.R | (color.G << 8) | (color.B << 16);

    private static Color FromBgr(object value)
    {
        var bgr = Convert.ToInt32(value);
        return Color.FromArgb(bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF);
    }

    /// <inheritdoc />
    public void CellFont(string sheetName, string cellRange, CellFontRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        var range = Target(Sheet(sheetName), cellRange);
        var font = range.Font;

        if (!string.IsNullOrWhiteSpace(request.FontName)) font.Name = request.FontName;
        if (request.FontSize > 0) font.Size = request.FontSize;

        // Every setting below means "leave it alone" in its None form.
        switch (request.FontStyle)
        {
            case FontStyleEnum.Regular: font.Bold = false; font.Italic = false; break;
            case FontStyleEnum.Bold: font.Bold = true; font.Italic = false; break;
            case FontStyleEnum.Italic: font.Bold = false; font.Italic = true; break;
            case FontStyleEnum.BoldItalic: font.Bold = true; font.Italic = true; break;
        }

        if (request.FontUnderLine != FontUnderLineEnum.None)
            font.Underline = Underline(request.FontUnderLine);

        switch (request.FontScript)
        {
            case FontScriptEnum.Superscript: font.Superscript = true; font.Subscript = false; break;
            case FontScriptEnum.Subscript: font.Superscript = false; font.Subscript = true; break;
            case FontScriptEnum.RemoveScript: font.Superscript = false; font.Subscript = false; break;
        }

        if (request.Strikethrough != SelectionYesNoNone.None)
            font.Strikethrough = request.Strikethrough == SelectionYesNoNone.Yes;

        if (!request.FontColor.IsEmpty) font.Color = Bgr(request.FontColor);
        if (!request.BackgroundColor.IsEmpty) range.Interior.Color = Bgr(request.BackgroundColor);

        Save();
    }

    private static Interop.XlUnderlineStyle Underline(FontUnderLineEnum style) => style switch
    {
        FontUnderLineEnum.Single => Interop.XlUnderlineStyle.xlUnderlineStyleSingle,
        FontUnderLineEnum.Double => Interop.XlUnderlineStyle.xlUnderlineStyleDouble,
        FontUnderLineEnum.SingleAccounting => Interop.XlUnderlineStyle.xlUnderlineStyleSingleAccounting,
        FontUnderLineEnum.DoubleAccounting => Interop.XlUnderlineStyle.xlUnderlineStyleDoubleAccounting,
        _ => Interop.XlUnderlineStyle.xlUnderlineStyleNone,
    };

    /// <inheritdoc />
    public CellFontRequest GetRangeStyle(string sheetName, string cellRange)
    {
        var range = Target(Sheet(sheetName), cellRange);
        var font = range.Font;

        return new CellFontRequest
        {
            FontName = Convert.ToString(font.Name) ?? string.Empty,
            FontSize = Convert.ToDouble(font.Size),
            FontColor = FromBgr(font.Color),
            BackgroundColor = FromBgr(range.Interior.Color),
        };
    }

    /// <inheritdoc />
    public void FillColor(string sheetName, string cellRange, Color background,
                          Color patternColor, PatternEnum pattern)
    {
        var interior = Target(Sheet(sheetName), cellRange).Interior;

        if (!background.IsEmpty) interior.Color = Bgr(background);
        if (pattern != PatternEnum.None)
        {
            interior.Pattern = (int)pattern;
            if (!patternColor.IsEmpty) interior.PatternColor = Bgr(patternColor);
        }
        Save();
    }

    /// <inheritdoc />
    public void ClearSheet(string sheetName, string cellRange, ClearOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        var range = Target(Sheet(sheetName), cellRange);

        if (options.All) range.Clear();
        if (options.Contents) range.ClearContents();
        if (options.Formats) range.ClearFormats();
        if (options.Comments) range.ClearComments();
        if (options.Notes) range.ClearNotes();
        if (options.Hyperlinks) range.Hyperlinks.Delete();
        if (options.Outline) range.ClearOutline();

        Save();
    }

    /// <inheritdoc />
    public void ColumnMove(string sheetName, string sourceColumns, string destinationColumns)
    {
        var sheet = Sheet(sheetName);
        sheet.Range[sourceColumns].EntireColumn.Cut(sheet.Range[destinationColumns].EntireColumn);
        Save();
    }

    /// <inheritdoc />
    public void CopyToClipboard(string sheetName, string cell, bool readFilter)
    {
        var range = Target(Sheet(sheetName), cell);

        // SpecialCells(xlCellTypeVisible) is what "only what the filter is showing" means
        // to Excel; without it a filtered copy brings the hidden rows along.
        if (readFilter) range = range.SpecialCells(Interop.XlCellType.xlCellTypeVisible);
        range.Copy();
    }

    /// <inheritdoc />
    public void PasteClipboard(string sheetName, string cell)
    {
        var sheet = Sheet(sheetName);
        sheet.Activate();
        sheet.Paste(sheet.Range[cell]);
        Save();
    }

    /// <inheritdoc />
    public void DeleteColumns(string sheetName, string[] columnsRange)
    {
        var sheet = Sheet(sheetName);

        // Right to left, so deleting one does not shift the ones still to go.
        foreach (var column in (columnsRange ?? []).Reverse())
            sheet.Range[column].EntireColumn.Delete();

        Save();
    }

    /// <inheritdoc />
    public void DeleteRows(string sheetName, string[] rowRange)
    {
        var sheet = Sheet(sheetName);

        // Bottom up, for the same reason as the columns above.
        foreach (var row in (rowRange ?? []).Reverse())
            sheet.Range[$"{row}:{row}".Replace(":", ":", StringComparison.Ordinal)].EntireRow.Delete();

        Save();
    }

    /// <inheritdoc />
    public (string[] Matches, int[] RowIndexes) Find(string sheetName, string cellRange,
                                                     string findText, FindReplaceEnum option, bool matchCase)
    {
        var range = Target(Sheet(sheetName), cellRange);
        var matches = new List<string>();
        var rows = new List<int>();

        var found = range.Find(
            findText,
            LookAt: option == FindReplaceEnum.Whole
                ? Interop.XlLookAt.xlWhole
                : Interop.XlLookAt.xlPart,
            MatchCase: matchCase);

        if (found is null) return ([], []);

        var first = found.Address[false, false];
        do
        {
            matches.Add(Convert.ToString(found.Text) ?? string.Empty);
            if (!rows.Contains(found.Row)) rows.Add(found.Row);

            found = range.FindNext(found);
        }
        // FindNext wraps round, so the walk stops when it comes back to where it started.
        while (found is not null && found.Address[false, false] != first);

        return ([.. matches], [.. rows]);
    }

    /// <inheritdoc />
    public bool FindReplace(string sheetName, string cellRange, string findText, string replaceText,
                            FindReplaceEnum option, bool matchCase)
    {
        var replaced = Target(Sheet(sheetName), cellRange).Replace(
            findText,
            replaceText,
            option == FindReplaceEnum.Whole ? Interop.XlLookAt.xlWhole : Interop.XlLookAt.xlPart,
            MatchCase: matchCase);

        Save();
        return Convert.ToBoolean(replaced);
    }

    /// <inheritdoc />
    public (int Index, string Name) FindLastColumn(string sheetName)
    {
        var used = Sheet(sheetName).UsedRange;
        var index = used.Column + used.Columns.Count - 1;
        return (index, ColumnName(index));
    }

    /// <inheritdoc />
    public int FindLastRow(string sheetName)
    {
        var used = Sheet(sheetName).UsedRange;
        return used.Row + used.Rows.Count - 1;
    }

    /// <inheritdoc />
    public bool FormatPainter(string sheetName, string cellRange,
                              string destinationSheet, string destinationRange)
    {
        var source = Sheet(sheetName).Range[cellRange];
        var destination = Sheet(string.IsNullOrWhiteSpace(destinationSheet) ? sheetName : destinationSheet)
            .Range[destinationRange];

        source.Copy();
        destination.PasteSpecial(Interop.XlPasteType.xlPasteFormats);
        _application.CutCopyMode = 0;

        Save();
        return true;
    }

    /// <inheritdoc />
    public bool IsMergedCell(string sheetName, string cell) =>
        Convert.ToBoolean(Sheet(sheetName).Range[cell].MergeCells);

    /// <inheritdoc />
    public void RefreshAll() => _workbook.RefreshAll();

    /// <inheritdoc />
    public void SelectCell(string sheetName, string cellRange)
    {
        var sheet = Sheet(sheetName);
        sheet.Activate();
        sheet.Range[cellRange].Select();
    }

    /// <inheritdoc />
    public void SetBorder(string sheetName, string cellRange, BorderEnum presets,
                          Interop.XlLineStyle lineStyle, double weight, Color color)
    {
        var range = Target(Sheet(sheetName), cellRange);

        if (presets == BorderEnum.NoBorder)
        {
            range.Borders.LineStyle = Interop.XlLineStyle.xlLineStyleNone;
            Save();
            return;
        }

        foreach (var edge in Edges(presets))
        {
            var border = range.Borders[edge];
            border.LineStyle = lineStyle;
            if (weight > 0) border.Weight = weight;
            if (!color.IsEmpty) border.Color = Bgr(color);
        }
        Save();
    }

    private static IEnumerable<Interop.XlBordersIndex> Edges(BorderEnum presets) => presets switch
    {
        BorderEnum.BottomBorder => [Interop.XlBordersIndex.xlEdgeBottom],
        BorderEnum.TopBorder => [Interop.XlBordersIndex.xlEdgeTop],
        BorderEnum.LeftBorder => [Interop.XlBordersIndex.xlEdgeLeft],
        BorderEnum.RightBorder => [Interop.XlBordersIndex.xlEdgeRight],
        BorderEnum.OutSideBorder =>
        [
            Interop.XlBordersIndex.xlEdgeLeft, Interop.XlBordersIndex.xlEdgeTop,
            Interop.XlBordersIndex.xlEdgeRight, Interop.XlBordersIndex.xlEdgeBottom,
        ],
        // AllBorder, which is the outside plus the two insides.
        _ =>
        [
            Interop.XlBordersIndex.xlEdgeLeft, Interop.XlBordersIndex.xlEdgeTop,
            Interop.XlBordersIndex.xlEdgeRight, Interop.XlBordersIndex.xlEdgeBottom,
            Interop.XlBordersIndex.xlInsideVertical, Interop.XlBordersIndex.xlInsideHorizontal,
        ],
    };

    /// <inheritdoc />
    public void PageSetup(string sheetName, PageSetupRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        var setup = Sheet(sheetName).PageSetup;

        if (request.LeftHeader.Length > 0) setup.LeftHeader = request.LeftHeader;
        if (request.CenterHeader.Length > 0) setup.CenterHeader = request.CenterHeader;
        if (request.RightHeader.Length > 0) setup.RightHeader = request.RightHeader;
        if (request.LeftFooter.Length > 0) setup.LeftFooter = request.LeftFooter;
        if (request.CenterFooter.Length > 0) setup.CenterFooter = request.CenterFooter;
        if (request.RightFooter.Length > 0) setup.RightFooter = request.RightFooter;

        if (request.LeftMargin > 0) setup.LeftMargin = request.LeftMargin;
        if (request.RightMargin > 0) setup.RightMargin = request.RightMargin;
        if (request.TopMargin > 0) setup.TopMargin = request.TopMargin;
        if (request.BottomMargin > 0) setup.BottomMargin = request.BottomMargin;
        if (request.HeaderMargin > 0) setup.HeaderMargin = request.HeaderMargin;
        if (request.FooterMargin > 0) setup.FooterMargin = request.FooterMargin;

        if (request.CenterOnPage != CenterOnPageEnum.None)
        {
            setup.CenterHorizontally = request.CenterOnPage == CenterOnPageEnum.Horizontally;
            setup.CenterVertically = request.CenterOnPage == CenterOnPageEnum.Vertically;
        }

        if (request.PageOrientation != PageOrientationEnum.None)
            setup.Orientation = request.PageOrientation == PageOrientationEnum.Portrait
                ? Interop.XlPageOrientation.xlPortrait
                : Interop.XlPageOrientation.xlLandscape;

        if (request.PageOrder != PageOrderEnum.None)
            setup.Order = request.PageOrder == PageOrderEnum.DownThenOver
                ? Interop.XlOrder.xlDownThenOver
                : Interop.XlOrder.xlOverThenDown;

        if (request.PageSize != PaperSizeEnum.PaperNone)
            setup.PaperSize = (Interop.XlPaperSize)(int)request.PageSize;

        if (request.FitToPagesWide > 0) setup.FitToPagesWide = request.FitToPagesWide;
        if (request.FitToPagesTall != TrueFaleNoneEnum.None)
            setup.FitToPagesTall = request.FitToPagesTall == TrueFaleNoneEnum.True ? 1 : false;

        if (request.PrintGridlines != TrueFaleNoneEnum.None)
            setup.PrintGridlines = request.PrintGridlines == TrueFaleNoneEnum.True;

        if (request.PrintZoom != TrueFaleNoneEnum.None)
            setup.Zoom = request.PrintZoom == TrueFaleNoneEnum.True;

        if (request.ZoomLevel > 0) setup.Zoom = request.ZoomLevel;

        Save();
    }
}
