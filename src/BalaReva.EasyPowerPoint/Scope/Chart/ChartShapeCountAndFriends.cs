using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;

namespace BalaReva.EasyPowerPoint.Scope.Chart;

/// <summary>Deletes a chart from a slide.</summary>
[DisplayName("Chart Delete")]
[Description("Deletes a chart from a slide.")]
public sealed class ChartDelete : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Which chart on the slide. Charts are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Chart Index")]
    [Description("Which chart on the slide. Charts are numbered from 1.")]
    public InArgument<int> ChartIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation) =>
        presentation.ChartDelete(SlideIndex.Get(context), ChartIndex.Get(context));
}

/// <summary>Copies a chart to the clipboard.</summary>
[DisplayName("Chart Copy To Clipboard")]
[Description("Copies a chart to the Windows clipboard.")]
public sealed class ChartCopyToClipboard : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Which chart on the slide. Charts are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Chart Index")]
    [Description("Which chart on the slide. Charts are numbered from 1.")]
    public InArgument<int> ChartIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation) =>
        presentation.ChartCopyToClipboard(SlideIndex.Get(context), ChartIndex.Get(context));
}

/// <summary>Moves and resizes a chart.</summary>
[DisplayName("Chart Format")]
[Description("Moves and resizes a chart on a slide.")]
public sealed class ChartFormat : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Which chart on the slide. Charts are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Chart Index")]
    [Description("Which chart on the slide. Charts are numbered from 1.")]
    public InArgument<int> ChartIndex { get; set; } = null!;

    /// <summary>Distance from the left edge of the slide, in points.</summary>
    [Category("Input")]
    [DisplayName("Chart Left")]
    [Description("Distance from the left edge of the slide, in points. Zero leaves it alone.")]
    public InArgument<float> ChartLeft { get; set; } = null!;

    /// <summary>Distance from the top edge of the slide, in points.</summary>
    [Category("Input")]
    [DisplayName("Chart Top")]
    [Description("Distance from the top edge of the slide, in points. Zero leaves it alone.")]
    public InArgument<float> ChartTop { get; set; } = null!;

    /// <summary>Width in points.</summary>
    [Category("Input")]
    [DisplayName("Chart Width")]
    [Description("Width in points. Zero leaves it alone.")]
    public InArgument<float> ChartWidth { get; set; } = null!;

    /// <summary>Height in points.</summary>
    [Category("Input")]
    [DisplayName("Chart Height")]
    [Description("Height in points. Zero leaves it alone.")]
    public InArgument<float> ChartHeight { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation) =>
        presentation.ChartFormat(
            SlideIndex.Get(context),
            ChartIndex.Get(context),
            ChartLeft?.Get(context) ?? 0,
            ChartTop?.Get(context) ?? 0,
            ChartWidth?.Get(context) ?? 0,
            ChartHeight?.Get(context) ?? 0);
}

/// <summary>Saves a slide's charts as images.</summary>
[DisplayName("Chart Image Extract")]
[Description("Saves each chart on a slide as an image file.")]
public sealed class ChartImageExtract : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Folder to save the images into.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Folder")]
    [Description("Folder to save the images into. Created if it does not exist.")]
    public InArgument<string> ImageFolder { get; set; } = null!;

    /// <summary>Image format to save in.</summary>
    [Category("Input")]
    [DisplayName("Image File Format")]
    [Description("Format to save the images in.")]
    public ImageFileFormatEnum ImageFileFormat { get; set; } = ImageFileFormatEnum.PNG;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation) =>
        presentation.ChartImageExtract(
            SlideIndex.Get(context),
            Require(context, ImageFolder, nameof(ImageFolder)),
            ImageFileFormat);
}
