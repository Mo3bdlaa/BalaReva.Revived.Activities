using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Tables;

/// <summary>Applies a named style and banding options to a table.</summary>
[DisplayName("Table Style")]
[Description("Applies a named Word table style and its banding options to a table.")]
public sealed class TableStyle : BaseNativeChild
{
    /// <summary>Which table. Tables are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table. Tables are numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <summary>Word table style to apply.</summary>
    [Category("Input")]
    [DisplayName("Style Name")]
    [Description("Name of a Word table style. Empty applies only the banding options.")]
    public InArgument<string> StyleName { get; set; } = null!;

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

    /// <summary>Style the first column.</summary>
    [Category("Input")]
    [DisplayName("First Column")]
    [Description("Style the first column.")]
    public bool FirstColumn { get; set; }

    /// <summary>Style the last column.</summary>
    [Category("Input")]
    [DisplayName("Last Column")]
    [Description("Style the last column.")]
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
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.TableStyle(
            TableIndex.Get(context),
            StyleName?.Get(context) ?? string.Empty,
            new WordTableStyleOptions
            {
                HeaderRow = HeaderRow,
                TotalRow = TotalRow,
                FirstColumn = FirstColumn,
                LastColumn = LastColumn,
                BandedRows = BandedRows,
                BandedColumns = BandedColumns,
            });
}
