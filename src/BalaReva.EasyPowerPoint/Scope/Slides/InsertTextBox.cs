using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Adds a text box to a slide.</summary>
[DisplayName("Insert Text Box")]
[Description("Adds a text box to a slide.")]
public sealed class InsertTextBox : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Text to put in the box.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Text")]
    [Description("Text to put in the box.")]
    public InArgument<string> Text { get; set; } = null!;

    /// <summary>Distance from the left edge, in points.</summary>
    [Category("Input")]
    [DisplayName("Text Box Left")]
    [Description("Distance from the left edge, in points.")]
    public InArgument<float> TextBoxLeft { get; set; } = null!;

    /// <summary>Distance from the top edge, in points.</summary>
    [Category("Input")]
    [DisplayName("Text Box Top")]
    [Description("Distance from the top edge, in points.")]
    public InArgument<float> TextBoxTop { get; set; } = null!;

    /// <summary>Width in points. Zero uses a default.</summary>
    [Category("Input")]
    [DisplayName("Text Box Width")]
    [Description("Width in points. Zero uses a default.")]
    public InArgument<float> TextBoxWidth { get; set; } = null!;

    /// <summary>Height in points. Zero uses a default.</summary>
    [Category("Input")]
    [DisplayName("Text Box Height")]
    [Description("Height in points. Zero uses a default.")]
    public InArgument<float> TextBoxHeight { get; set; } = null!;

    /// <summary>Font name. Empty leaves the default.</summary>
    [Category("Input")]
    [DisplayName("Font Name")]
    [Description("Font name. Empty leaves the default.")]
    public InArgument<string> FontName { get; set; } = null!;

    /// <summary>Font size in points. Zero leaves the default.</summary>
    [Category("Input")]
    [DisplayName("Font Size")]
    [Description("Font size in points. Zero leaves the default.")]
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

    /// <summary>How the text sits inside the box.</summary>
    [Category("Input")]
    [DisplayName("Text Alignment")]
    [Description("How the text sits inside the box.")]
    public EnumTextEffectAlignment TextAlignment { get; set; } = EnumTextEffectAlignment.Left;

    /// <summary>Which way the text runs.</summary>
    [Category("Input")]
    [DisplayName("Text Orientation")]
    [Description("Which way the text runs.")]
    public TextOrientationEnum TextOrientation { get; set; } = TextOrientationEnum.Horizontal;

    /// <summary>Where the box sits in the slide stacking order.</summary>
    [Category("Input")]
    [DisplayName("Z Order")]
    [Description("Where the box sits in the slide stacking order.")]
    public ZOrderCmdEnum ZOrder { get; set; } = ZOrderCmdEnum.BringToFront;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.InsertTextBox(SlideIndex.Get(context), new TextBoxRequest
        {
            Text = Require(context, Text, nameof(Text)),
            Left = TextBoxLeft?.Get(context) ?? 0,
            Top = TextBoxTop?.Get(context) ?? 0,
            Width = TextBoxWidth?.Get(context) ?? 0,
            Height = TextBoxHeight?.Get(context) ?? 0,
            Alignment = TextAlignment,
            Orientation = TextOrientation,
            ZOrder = ZOrder,
            Style = new TextStyleRequest
            {
                FontName = FontName?.Get(context) ?? string.Empty,
                FontSize = FontSize?.Get(context) ?? 0,
                Bold = FontBold,
                Italic = FontItalic,
            },
        });
}
