using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Sets a table's banding and emphasis options.</summary>
[DisplayName("Style Option")]
[Description("Sets a table's banding and emphasis options.")]
public sealed class StyleOption : BaseTableNativeChild
{
    /// <summary>Style the first row as a header. None leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Header Row")]
    [Description("Style the first row as a header. None leaves it alone.")]
    public TrueFalseNoneEnum HeaderRow { get; set; } = TrueFalseNoneEnum.None;

    /// <summary>Style the last row as a totals row. None leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Total Row")]
    [Description("Style the last row as a totals row. None leaves it alone.")]
    public TrueFalseNoneEnum TotalRow { get; set; } = TrueFalseNoneEnum.None;

    /// <summary>Emphasise the first column. None leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("First Column")]
    [Description("Emphasise the first column. None leaves it alone.")]
    public TrueFalseNoneEnum FirstColumn { get; set; } = TrueFalseNoneEnum.None;

    /// <summary>Emphasise the last column. None leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Last Column")]
    [Description("Emphasise the last column. None leaves it alone.")]
    public TrueFalseNoneEnum LastColumn { get; set; } = TrueFalseNoneEnum.None;

    /// <summary>Band the rows. None leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Banded Rows")]
    [Description("Band the rows. None leaves it alone.")]
    public TrueFalseNoneEnum BandedRows { get; set; } = TrueFalseNoneEnum.None;

    /// <summary>Band the columns. None leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Banded Columns")]
    [Description("Band the columns. None leaves it alone.")]
    public TrueFalseNoneEnum BandedColumns { get; set; } = TrueFalseNoneEnum.None;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.StyleOption(Table(context), new TableStyleOptions
        {
            HeaderRow = HeaderRow,
            TotalRow = TotalRow,
            FirstColumn = FirstColumn,
            LastColumn = LastColumn,
            BandedRows = BandedRows,
            BandedColumns = BandedColumns,
        });
}
