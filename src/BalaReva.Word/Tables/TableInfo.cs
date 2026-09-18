using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Tables;

/// <summary>Reports the shape and style flags of one table.</summary>
[DisplayName("Table Info")]
[Description("Reports the row and column counts and style flags of one table.")]
public sealed class TableInfo : BaseNativeChild
{
    /// <summary>Which table. Tables are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table to report on. Tables are numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <summary>Number of rows.</summary>
    [Category("Output")]
    [DisplayName("Total Rows")]
    [Description("Number of rows in the table.")]
    public OutArgument<int> TotalRows { get; set; } = null!;

    /// <summary>Number of columns.</summary>
    [Category("Output")]
    [DisplayName("Total Columns")]
    [Description("Number of columns in the table.")]
    public OutArgument<int> TotalColumns { get; set; } = null!;

    /// <summary>Whether the style marks a header row.</summary>
    [Category("Output")]
    [DisplayName("Has Header Row")]
    [Description("Whether the table's style marks a header row.")]
    public OutArgument<bool> HasHeaderRow { get; set; } = null!;

    /// <summary>Whether the style bands rows.</summary>
    [Category("Output")]
    [DisplayName("Has Banded Rows")]
    [Description("Whether the table's style bands its rows.")]
    public OutArgument<bool> HasBandedRows { get; set; } = null!;

    /// <summary>Whether the style bands columns.</summary>
    [Category("Output")]
    [DisplayName("Has Banded Columns")]
    [Description("Whether the table's style bands its columns.")]
    public OutArgument<bool> HasBandedColumns { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document)
    {
        var info = document.TableInfo(TableIndex.Get(context));
        TotalRows.Set(context, info.TotalRows);
        TotalColumns.Set(context, info.TotalColumns);
        HasHeaderRow.Set(context, info.HasHeaderRow);
        HasBandedRows.Set(context, info.HasBandedRows);
        HasBandedColumns.Set(context, info.HasBandedColumns);
    }
}
