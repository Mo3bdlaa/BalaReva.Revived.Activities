using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Saves one slide out as its own presentation.</summary>
[DisplayName("Slide Extractor")]
[Description("Saves one slide out as its own presentation.")]
public sealed class SlideExtractor : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Full path of the file that was written.</summary>
    [Category("Output")]
    [DisplayName("Slide Result")]
    [Description("Full path of the file that was written.")]
    public OutArgument<string> SlideResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => SlideResult.Set(context, presentation.SlideExtractor(SlideIndex.Get(context)));
}
