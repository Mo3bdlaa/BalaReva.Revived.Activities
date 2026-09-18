using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Places an image on a slide.</summary>
[DisplayName("Insert Picture")]
[Description("Places an image on a slide.")]
public sealed class InsertPicture : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Full path of the image to place.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Path")]
    [Description("Full path of the image to place.")]
    public InArgument<string> ImagePath { get; set; } = null!;

    /// <summary>Distance from the left edge, in points.</summary>
    [Category("Input")]
    [DisplayName("Image Left")]
    [Description("Distance from the left edge, in points.")]
    public InArgument<float> ImageLeft { get; set; } = null!;

    /// <summary>Distance from the top edge, in points.</summary>
    [Category("Input")]
    [DisplayName("Image Top")]
    [Description("Distance from the top edge, in points.")]
    public InArgument<float> ImageTop { get; set; } = null!;

    /// <summary>Width in points. Zero keeps the image own width.</summary>
    [Category("Input")]
    [DisplayName("Image Width")]
    [Description("Width in points. Zero keeps the image own width.")]
    public InArgument<float> ImageWidth { get; set; } = null!;

    /// <summary>Height in points. Zero keeps the image own height.</summary>
    [Category("Input")]
    [DisplayName("Image Height")]
    [Description("Height in points. Zero keeps the image own height.")]
    public InArgument<float> ImageHeight { get; set; } = null!;

    /// <summary>Where the image sits in the slide stacking order.</summary>
    [Category("Input")]
    [DisplayName("Z Order")]
    [Description("Where the image sits in the slide stacking order.")]
    public ZOrderCmdEnum ZOrder { get; set; } = ZOrderCmdEnum.BringToFront;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.InsertPicture(SlideIndex.Get(context), new PictureRequest
        {
            ImagePath = Require(context, ImagePath, nameof(ImagePath)),
            Left = ImageLeft?.Get(context) ?? 0,
            Top = ImageTop?.Get(context) ?? 0,
            Width = ImageWidth?.Get(context) ?? 0,
            Height = ImageHeight?.Get(context) ?? 0,
            ZOrder = ZOrder,
        });
}
