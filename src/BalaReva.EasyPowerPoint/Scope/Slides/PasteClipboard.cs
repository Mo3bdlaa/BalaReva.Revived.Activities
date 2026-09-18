using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Pastes the clipboard onto a slide.</summary>
[DisplayName("Paste Clipboard")]
[Description("Pastes the clipboard onto a slide.")]
public sealed class PasteClipboard : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Distance from the left edge, in points. Zero leaves where it lands.</summary>
    [Category("Input")]
    [DisplayName("Object Left")]
    [Description("Distance from the left edge, in points. Zero leaves where it lands.")]
    public InArgument<double> ObjectLeft { get; set; } = null!;

    /// <summary>Distance from the top edge, in points. Zero leaves where it lands.</summary>
    [Category("Input")]
    [DisplayName("Object Top")]
    [Description("Distance from the top edge, in points. Zero leaves where it lands.")]
    public InArgument<double> ObjectTop { get; set; } = null!;

    /// <summary>Width in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Object Width")]
    [Description("Width in points. Zero leaves it alone.")]
    public InArgument<double> ObjectWidth { get; set; } = null!;

    /// <summary>Height in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Object Height")]
    [Description("Height in points. Zero leaves it alone.")]
    public InArgument<double> ObjectHeight { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.PasteClipboard(
            SlideIndex.Get(context),
            ObjectLeft?.Get(context) ?? 0,
            ObjectTop?.Get(context) ?? 0,
            ObjectWidth?.Get(context) ?? 0,
            ObjectHeight?.Get(context) ?? 0);
}
