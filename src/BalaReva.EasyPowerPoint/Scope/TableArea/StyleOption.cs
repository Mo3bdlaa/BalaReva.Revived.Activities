using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Sets a table\u0027s banding and emphasis options.</summary>
[DisplayName("Style Option")]
[Description("Sets a table\u0027s banding and emphasis options.")]
public sealed class StyleOption : BaseTableNativeChild
{
    /// <summary>Style the first row as a header.</summary>
    [Category("Input")]
    [DisplayName("Header Row")]
    [Description("Style the first row as a header.")]
    public bool HeaderRow { get; set; }

    /// <summary>Style the last row as a totals row.</summary>
    [Category("Input")]
    [DisplayName("Total Row")]
    [Description("Style the last row as a totals row.")]
    public bool TotalRow { get; set; }

    /// <summary>Emphasise the first column.</summary>
    [Category("Input")]
    [DisplayName("First Column")]
    [Description("Emphasise the first column.")]
    public bool FirstColumn { get; set; }

    /// <summary>Emphasise the last column.</summary>
    [Category("Input")]
    [DisplayName("Last Column")]
    [Description("Emphasise the last column.")]
    public bool LastColumn { get; set; }

    /// <summary>Band the rows.</summary>
    [Category("Input")]
    [DisplayName("Banded Rows")]
    [Description("Band the rows.")]
    public bool BandedRows { get; set; }

    /// <summary>Band the columns.</summary>
    [Category("Input")]
    [DisplayName("Banded Columns")]
    [Description("Band the columns.")]
    public bool BandedColumns { get; set; }

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
