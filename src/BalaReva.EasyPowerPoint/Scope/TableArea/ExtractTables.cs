using System.Activities;
using System.ComponentModel;
using System.Data;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Reads every table on a slide.</summary>
[DisplayName("Extract Tables")]
[Description("Reads every table on a slide.")]
public sealed class ExtractTables : BaseSlideNativeChild
{
    /// <summary>One DataTable per table on the slide.</summary>
    [Category("Output")]
    [DisplayName("Output Tables")]
    [Description("One DataTable per table on the slide.")]
    public OutArgument<DataTable[]> OutputTables { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => OutputTables.Set(
            context, presentation.ExtractTables([SlideIndex.Get(context)], hasHeader: false));
}
