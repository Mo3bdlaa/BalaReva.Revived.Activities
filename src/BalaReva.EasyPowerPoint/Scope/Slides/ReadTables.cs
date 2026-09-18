using System.Activities;
using System.ComponentModel;
using System.Data;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Reads every table on a slide.</summary>
[DisplayName("Read Tables")]
[Description("Reads every table on a slide.")]
public sealed class ReadTables : BaseNativeChild
{
    /// <summary>Slides to read, numbered from 1. Empty means every slide.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Slides to read, numbered from 1. Empty means every slide.")]
    public InArgument<int[]> SlideIndex { get; set; } = null!;

    /// <summary>Treat the first row of each table as its column names.</summary>
    [Category("Input")]
    [DisplayName("Has Header")]
    [Description("Treat the first row of each table as its column names.")]
    public bool HasHeader { get; set; }

    /// <summary>One DataTable per table on the slide.</summary>
    [Category("Output")]
    [DisplayName("Result Set")]
    [Description("One DataTable per table on the slide.")]
    public OutArgument<DataSet> ResultSet { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
    {
        var set = new DataSet();
        set.Tables.AddRange(presentation.ExtractTables(SlideIndex?.Get(context) ?? [], HasHeader));
        ResultSet.Set(context, set);
    }
}
