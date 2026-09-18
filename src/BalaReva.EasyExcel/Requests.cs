using System.Data;
using System.Drawing;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel;

// None of the types in this file are part of the published binding surface. They exist
// so the service signatures stay readable: several of these activities carry a dozen or
// more arguments, and PageSetup carries twenty-one.

/// <summary>How an <c>ExcelScope</c> should open its workbook.</summary>
public sealed class ExcelOpenRequest
{
    /// <summary>Full path of the workbook.</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>Password needed to open it.</summary>
    public string FilePassword { get; set; } = string.Empty;

    /// <summary>Password needed to change it.</summary>
    public string ModifyPassword { get; set; } = string.Empty;

    /// <summary>Let Excel show its own prompts and warnings.</summary>
    public bool DisplayAlerts { get; set; }

    /// <summary>Show the Excel window.</summary>
    public bool Visible { get; set; }

    /// <summary>Whether macros in the workbook may run.</summary>
    public bool MacrosEnabled { get; set; }

    /// <summary>Whether links to other workbooks are refreshed on open.</summary>
    public AutomaticLinkEnum UpdateAutoLinks { get; set; } = AutomaticLinkEnum.Default;
}

/// <summary>Which sheet of which workbook <c>SaveAsSheet</c> should write out.</summary>
public sealed class SaveSheetRequest
{
    /// <summary>Full path of the workbook to read.</summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>Password needed to open it.</summary>
    public string FilePassword { get; set; } = string.Empty;

    /// <summary>Password needed to change it.</summary>
    public string ModifyPassword { get; set; } = string.Empty;

    /// <summary>Full path of the file to write.</summary>
    public string NewFileName { get; set; } = string.Empty;

    /// <summary>Sheet to write out. Empty means the active sheet.</summary>
    public string Sheet { get; set; } = string.Empty;
}

/// <summary>Which chart on which sheet, by name or by position.</summary>
/// <remarks>The name wins when both are given, which is how the published package behaved.</remarks>
public sealed class ChartRef
{
    /// <summary>Sheet the chart is on. Empty means the active sheet.</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Name of the chart shape.</summary>
    public string ChartName { get; set; } = string.Empty;

    /// <summary>Position of the chart on the sheet, numbered from 1.</summary>
    public int ChartIndex { get; set; }
}

/// <summary>Which picture on which sheet, by name or by position.</summary>
public sealed class ImageRef
{
    /// <summary>Sheet the picture is on. Empty means the active sheet.</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Name of the picture shape.</summary>
    public string ImageName { get; set; } = string.Empty;

    /// <summary>Position of the picture on the sheet, numbered from 1.</summary>
    public int ImageIndex { get; set; }
}

/// <summary>Which table on a sheet, by name or by position.</summary>
public sealed class TableRef
{
    /// <summary>Name of the table.</summary>
    public string TableName { get; set; } = string.Empty;

    /// <summary>Position of the table on the sheet, numbered from 1.</summary>
    public int TableIndex { get; set; }
}

/// <summary>Where a shape sits and how big it is, in points. Zero leaves one alone.</summary>
public sealed class ChartBounds
{
    /// <summary>Distance from the left edge of the sheet.</summary>
    public double Left { get; set; }

    /// <summary>Distance from the top edge of the sheet.</summary>
    public double Top { get; set; }

    /// <summary>Width.</summary>
    public double Width { get; set; }

    /// <summary>Height.</summary>
    public double Height { get; set; }
}

/// <summary>Which of Excel's range functions to evaluate.</summary>
public enum RangeFunctionKind
{
    /// <summary>Arithmetic mean of the numbers in the range.</summary>
    Average,

    /// <summary>How many cells in the range hold numbers.</summary>
    Count,

    /// <summary>How many cells in the range are not empty.</summary>
    CountA,

    /// <summary>Largest number in the range.</summary>
    Max,

    /// <summary>Smallest number in the range.</summary>
    Min,

    /// <summary>Sum of the numbers in the range.</summary>
    Sum,
}

/// <summary>Font and fill for a range. An unset member leaves that setting alone.</summary>
public sealed class CellFontRequest
{
    /// <summary>Font name.</summary>
    public string FontName { get; set; } = string.Empty;

    /// <summary>Font size in points. Zero leaves it alone.</summary>
    public double FontSize { get; set; }

    /// <summary>Bold, italic, both, or neither.</summary>
    public FontStyleEnum FontStyle { get; set; } = FontStyleEnum.None;

    /// <summary>Underline style.</summary>
    public FontUnderLineEnum FontUnderLine { get; set; } = FontUnderLineEnum.None;

    /// <summary>Superscript or subscript.</summary>
    public FontScriptEnum FontScript { get; set; } = FontScriptEnum.None;

    /// <summary>Strikethrough.</summary>
    public SelectionYesNoNone Strikethrough { get; set; } = SelectionYesNoNone.None;

    /// <summary>Font colour. Empty leaves it alone.</summary>
    public Color FontColor { get; set; }

    /// <summary>Fill colour. Empty leaves it alone.</summary>
    public Color BackgroundColor { get; set; }
}

/// <summary>Which parts of a range <c>ClearSheet</c> should clear.</summary>
public sealed class ClearOptions
{
    /// <summary>Clear everything.</summary>
    public bool All { get; set; }

    /// <summary>Clear the values and formulas.</summary>
    public bool Contents { get; set; }

    /// <summary>Clear the formatting.</summary>
    public bool Formats { get; set; }

    /// <summary>Clear the threaded comments.</summary>
    public bool Comments { get; set; }

    /// <summary>Clear the notes.</summary>
    public bool Notes { get; set; }

    /// <summary>Clear the hyperlinks.</summary>
    public bool Hyperlinks { get; set; }

    /// <summary>Clear the outline grouping.</summary>
    public bool Outline { get; set; }
}

/// <summary>Everything <c>PageSetup</c> can set on a sheet.</summary>
public sealed class PageSetupRequest
{
    /// <summary>Header text, left, centre and right.</summary>
    public string LeftHeader { get; set; } = string.Empty;

    /// <summary>Header text, centre.</summary>
    public string CenterHeader { get; set; } = string.Empty;

    /// <summary>Header text, right.</summary>
    public string RightHeader { get; set; } = string.Empty;

    /// <summary>Footer text, left.</summary>
    public string LeftFooter { get; set; } = string.Empty;

    /// <summary>Footer text, centre.</summary>
    public string CenterFooter { get; set; } = string.Empty;

    /// <summary>Footer text, right.</summary>
    public string RightFooter { get; set; } = string.Empty;

    /// <summary>Left margin in points.</summary>
    public int LeftMargin { get; set; }

    /// <summary>Right margin in points.</summary>
    public int RightMargin { get; set; }

    /// <summary>Top margin in points.</summary>
    public int TopMargin { get; set; }

    /// <summary>Bottom margin in points.</summary>
    public int BottomMargin { get; set; }

    /// <summary>Header margin in points.</summary>
    public int HeaderMargin { get; set; }

    /// <summary>Footer margin in points.</summary>
    public int FooterMargin { get; set; }

    /// <summary>Whether the printout is centred on the page.</summary>
    public CenterOnPageEnum CenterOnPage { get; set; } = CenterOnPageEnum.None;

    /// <summary>Portrait or landscape.</summary>
    public PageOrientationEnum PageOrientation { get; set; } = PageOrientationEnum.None;

    /// <summary>Down then over, or over then down.</summary>
    public PageOrderEnum PageOrder { get; set; } = PageOrderEnum.None;

    /// <summary>Paper size.</summary>
    public PaperSizeEnum PageSize { get; set; } = PaperSizeEnum.PaperNone;

    /// <summary>How many pages wide to fit to. Zero leaves it alone.</summary>
    public int FitToPagesWide { get; set; }

    /// <summary>Whether to fit to a page height at all.</summary>
    public TrueFaleNoneEnum FitToPagesTall { get; set; } = TrueFaleNoneEnum.None;

    /// <summary>Print the gridlines.</summary>
    public TrueFaleNoneEnum PrintGridlines { get; set; } = TrueFaleNoneEnum.None;

    /// <summary>Scale the printout rather than fitting it to pages.</summary>
    public TrueFaleNoneEnum PrintZoom { get; set; } = TrueFaleNoneEnum.None;

    /// <summary>Zoom percentage. Zero leaves it alone.</summary>
    public int ZoomLevel { get; set; }
}

/// <summary>Which rows <c>MergeSheetByRow</c> should append, and from where.</summary>
public sealed class MergeRequest
{
    /// <summary>Full path of the workbook to append from.</summary>
    public string AppendFileName { get; set; } = string.Empty;

    /// <summary>Password needed to open it.</summary>
    public string AppendFilePassword { get; set; } = string.Empty;

    /// <summary>Password needed to change it.</summary>
    public string AppendModifyPassword { get; set; } = string.Empty;

    /// <summary>Sheet to append from.</summary>
    public string AppendSheet { get; set; } = string.Empty;

    /// <summary>Range to append. Empty means everything with anything in it.</summary>
    public string AppendCellRange { get; set; } = string.Empty;

    /// <summary>Column to start writing at, for example A.</summary>
    public string StartColumn { get; set; } = string.Empty;
}
