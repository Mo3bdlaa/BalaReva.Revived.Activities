using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Tables;

/// <summary>Writes a cell's text and font.</summary>
[DisplayName("Set Table Value")]
[Description("Writes text into a table cell and sets its font.")]
public sealed class SetTableValue : BaseNativeChild
{
    /// <summary>Which table. Tables are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table. Tables are numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <summary>Which row. Rows are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Row Index")]
    [Description("Which row. Rows are numbered from 1.")]
    public InArgument<int> RowIndex { get; set; } = null!;

    /// <summary>Which column. Columns are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Column Index")]
    [Description("Which column. Columns are numbered from 1.")]
    public InArgument<int> ColumnIndex { get; set; } = null!;

    /// <summary>Text to write.</summary>
    [Category("Input")]
    [DisplayName("Text Value")]
    [Description("Text to write into the cell.")]
    public InArgument<string> TextValue { get; set; } = null!;

    /// <summary>Font name.</summary>
    [Category("Input")]
    [DisplayName("Text Font Name")]
    [Description("Font name. Empty leaves the cell's font alone.")]
    public InArgument<string> TextFontName { get; set; } = null!;

    /// <summary>Font size in points.</summary>
    [Category("Input")]
    [DisplayName("Font Size")]
    [Description("Font size in points. Zero leaves it alone.")]
    public InArgument<float> FontSize { get; set; } = null!;

    /// <summary>Bold. Select leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Bold")]
    [Description("Bold. Select leaves the cell's setting alone.")]
    public EnumSelectBoolean FontBold { get; set; } = EnumSelectBoolean.Select;

    /// <summary>Italic. Select leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Italic")]
    [Description("Italic. Select leaves the cell's setting alone.")]
    public EnumSelectBoolean FontItalic { get; set; } = EnumSelectBoolean.Select;

    /// <summary>Underline. Select leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Underline")]
    [Description("Underline. Select leaves the cell's setting alone.")]
    public EnumSelectBoolean FontUnderline { get; set; } = EnumSelectBoolean.Select;

    /// <summary>Strikethrough. Select leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Strikeout")]
    [Description("Strikethrough. Select leaves the cell's setting alone.")]
    public EnumSelectBoolean FontStrikeout { get; set; } = EnumSelectBoolean.Select;

    /// <summary>Superscript. Select leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Superscript")]
    [Description("Superscript. Select leaves the cell's setting alone.")]
    public EnumSelectBoolean FontSuperscript { get; set; } = EnumSelectBoolean.Select;

    /// <summary>Subscript. Select leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Subscript")]
    [Description("Subscript. Select leaves the cell's setting alone.")]
    public EnumSelectBoolean FontSubscript { get; set; } = EnumSelectBoolean.Select;

    /// <summary>Vertical alignment in the cell.</summary>
    [Category("Input")]
    [DisplayName("Text Vertical Alignment")]
    [Description("Vertical alignment within the cell. Select leaves it alone.")]
    public EnumCellVerticalAlignment TextVerticalAlignment { get; set; } = EnumCellVerticalAlignment.Select;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.SetTableValue(
            TableIndex.Get(context),
            RowIndex.Get(context),
            ColumnIndex.Get(context),
            new WordCellFormat
            {
                TextValue = TextValue?.Get(context) ?? string.Empty,
                FontName = TextFontName?.Get(context) ?? string.Empty,
                FontSize = FontSize?.Get(context) ?? 0,
                Bold = FontBold,
                Italic = FontItalic,
                Underline = FontUnderline,
                Strikeout = FontStrikeout,
                Superscript = FontSuperscript,
                Subscript = FontSubscript,
                VerticalAlignment = TextVerticalAlignment,
            });
}
