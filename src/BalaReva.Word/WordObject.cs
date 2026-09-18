namespace BalaReva.Word;

/// <summary>
/// The document a <c>WordScope</c> is open on, handed to its body.
/// </summary>
/// <remarks>
/// Note <c>ModiPassword</c> rather than ModifyPassword: that is how the published
/// package spelled it, and a workflow binds by property name.
/// </remarks>
public sealed class WordObject
{
    /// <summary>Full path of the open document.</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>Password used to open the document.</summary>
    public string Password { get; set; } = string.Empty;

    /// <summary>Password used to permit changes to the document.</summary>
    public string ModiPassword { get; set; } = string.Empty;
}

/// <summary>Counts that <c>WordStatistics</c> reports.</summary>
/// <remarks>Not part of the published surface; it keeps the service to one return value.</remarks>
public sealed class WordStatisticsResult
{
    /// <summary>Characters, not counting spaces.</summary>
    public long Characters { get; set; }

    /// <summary>Characters, counting spaces.</summary>
    public long CharactersWithSpaces { get; set; }

    /// <summary>Lines.</summary>
    public long Lines { get; set; }

    /// <summary>Pages.</summary>
    public long Pages { get; set; }

    /// <summary>Paragraphs.</summary>
    public long Paragraphs { get; set; }

    /// <summary>Words.</summary>
    public long WordCount { get; set; }
}

/// <summary>What <c>TableInfo</c> reports about one table.</summary>
/// <remarks>Not part of the published surface; it keeps the service to one return value.</remarks>
public sealed class WordTableInfo
{
    /// <summary>Number of rows.</summary>
    public int TotalRows { get; set; }

    /// <summary>Number of columns.</summary>
    public int TotalColumns { get; set; }

    /// <summary>Whether the table's style marks a header row.</summary>
    public bool HasHeaderRow { get; set; }

    /// <summary>Whether the table's style bands its rows.</summary>
    public bool HasBandedRows { get; set; }

    /// <summary>Whether the table's style bands its columns.</summary>
    public bool HasBandedColumns { get; set; }
}

/// <summary>Formatting applied by <c>InsertHeaderFooter</c>.</summary>
/// <remarks>Not part of the published surface; it keeps the service signature readable.</remarks>
public sealed class HeaderFooterText
{
    /// <summary>Text for odd pages, and for every page when the others are unset.</summary>
    public string OddPageText { get; set; } = string.Empty;

    /// <summary>Text for even pages.</summary>
    public string EvenPageText { get; set; } = string.Empty;

    /// <summary>Text for the first page.</summary>
    public string FirstPageText { get; set; } = string.Empty;

    /// <summary>Font name.</summary>
    public string FontName { get; set; } = string.Empty;

    /// <summary>Font size in points. Zero leaves it alone.</summary>
    public float FontSize { get; set; }

    /// <summary>Bold.</summary>
    public Utilities.EnumBoolean Bold { get; set; } = Utilities.EnumBoolean.False;

    /// <summary>Italic.</summary>
    public Utilities.EnumBoolean Italic { get; set; } = Utilities.EnumBoolean.False;

    /// <summary>Underline.</summary>
    public Utilities.EnumBoolean Underline { get; set; } = Utilities.EnumBoolean.False;

    /// <summary>Strikethrough.</summary>
    public Utilities.EnumBoolean Strikeout { get; set; } = Utilities.EnumBoolean.False;

    /// <summary>Horizontal alignment.</summary>
    public Utilities.HeaderFooterParagraphAlignment Alignment { get; set; } =
        Utilities.HeaderFooterParagraphAlignment.Select;
}

/// <summary>Images placed by <c>InsertHeaderFooterImage</c>.</summary>
/// <remarks>Not part of the published surface; it keeps the service signature readable.</remarks>
public sealed class HeaderFooterImages
{
    /// <summary>Image for odd pages.</summary>
    public string OddPageImage { get; set; } = string.Empty;

    /// <summary>Image for even pages.</summary>
    public string EvenPageImage { get; set; } = string.Empty;

    /// <summary>Image for the first page.</summary>
    public string FirstPageImage { get; set; } = string.Empty;

    /// <summary>Horizontal alignment.</summary>
    public Utilities.HeaderFooterParagraphAlignment Alignment { get; set; } =
        Utilities.HeaderFooterParagraphAlignment.Select;
}
