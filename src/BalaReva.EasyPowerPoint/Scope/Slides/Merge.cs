using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Appends slides from another presentation.</summary>
[DisplayName("Merge")]
[Description("Appends slides from another presentation.")]
public sealed class Merge : BaseNativeChild
{
    /// <summary>Full path of the presentation to take slides from.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Source Ppt")]
    [Description("Full path of the presentation to take slides from.")]
    public InArgument<string> SourcePpt { get; set; } = null!;

    /// <summary>First slide to take. Zero takes all of them.</summary>
    [Category("Input")]
    [DisplayName("Start Slide Index")]
    [Description("First slide to take. Zero takes all of them.")]
    public InArgument<int> StartSlideIndex { get; set; } = null!;

    /// <summary>Last slide to take. Zero takes all of them.</summary>
    [Category("Input")]
    [DisplayName("End Slide Index")]
    [Description("Last slide to take. Zero takes all of them.")]
    public InArgument<int> EndSlideIndex { get; set; } = null!;

    /// <summary>Slide in this presentation to insert after. Zero inserts at the start.</summary>
    [Category("Input")]
    [DisplayName("Slide After")]
    [Description("Slide in this presentation to insert after. Zero inserts at the start.")]
    public InArgument<int> SlideAfter { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.Merge(
            Require(context, SourcePpt, nameof(SourcePpt)),
            StartSlideIndex?.Get(context) ?? 0,
            EndSlideIndex?.Get(context) ?? 0,
            SlideAfter?.Get(context) ?? 0);
}
