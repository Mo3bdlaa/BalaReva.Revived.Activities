using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Sets the font on every cell of a table.</summary>
[DisplayName("Font Option")]
[Description("Sets the font on every cell of a table.")]
public sealed class FontOption : BaseTableNativeChild
{
    /// <summary>Font name. Empty leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Name")]
    [Description("Font name. Empty leaves it alone.")]
    public InArgument<string> FontName { get; set; } = null!;

    /// <summary>Font size in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Size")]
    [Description("Font size in points. Zero leaves it alone.")]
    public InArgument<float> FontSize { get; set; } = null!;

    /// <summary>Bold. None leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Bold")]
    [Description("Bold. None leaves it alone.")]
    public TrueFalseNoneEnum FontBold { get; set; } = TrueFalseNoneEnum.None;

    /// <summary>Italic. None leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Italic")]
    [Description("Italic. None leaves it alone.")]
    public TrueFalseNoneEnum FontItalic { get; set; } = TrueFalseNoneEnum.None;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.FontOption(Table(context), new TextStyleRequest
        {
            FontName = FontName?.Get(context) ?? string.Empty,
            FontSize = FontSize?.Get(context) ?? 0,
            Bold = FontBold,
            Italic = FontItalic,
        });
}
