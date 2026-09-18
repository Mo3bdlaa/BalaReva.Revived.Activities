using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Lists the names of a slide\u0027s tables.</summary>
[DisplayName("Get Table Names")]
[Description("Lists the names of a slide\u0027s tables.")]
public sealed class GetTableNames : BaseSlideNativeChild
{
    /// <summary>Names of the table shapes on the slide.</summary>
    [Category("Output")]
    [DisplayName("Output Result")]
    [Description("Names of the table shapes on the slide.")]
    public OutArgument<string[]> OutputResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => OutputResult.Set(context, presentation.GetTableNames(SlideIndex.Get(context)));
}
