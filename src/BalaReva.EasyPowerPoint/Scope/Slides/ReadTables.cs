using System.Activities;
using System.ComponentModel;
using System.Data;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Reads every table on a slide.</summary>
/// <remarks>
/// HasHeader is accepted for compatibility but does not change the result: the
/// service reads every row as data, so a header row comes back as the first row.
/// See docs/REVIVAL.md.
/// </remarks>
[DisplayName("Read Tables")]
[Description("Reads every table on a slide.")]
public sealed class ReadTables : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Treat the first row of each table as its column names.</summary>
    [Category("Input")]
    [DisplayName("Has Header")]
    [Description("Treat the first row of each table as its column names.")]
    public InArgument<bool> HasHeader { get; set; } = null!;

    /// <summary>One DataTable per table on the slide.</summary>
    [Category("Output")]
    [DisplayName("Result Set")]
    [Description("One DataTable per table on the slide.")]
    public OutArgument<DataSet> ResultSet { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => ResultSet.Set(context, presentation.ExtractTables(SlideIndex.Get(context)));
}
