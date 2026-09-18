using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Refreshes the data behind a slide charts.</summary>
[DisplayName("Refresh Data")]
[Description("Refreshes the data behind a slide charts.")]
public sealed class RefreshData : BaseNativeChild
{
    /// <summary>Slides to refresh, numbered from 1. Empty means every slide.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Slides to refresh, numbered from 1. Empty means every slide.")]
    public InArgument<short[]> SlideIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.RefreshData(SlideIndex?.Get(context) ?? []);
}
