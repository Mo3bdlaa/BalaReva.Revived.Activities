using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Tables;

/// <summary>Adds rows to a table.</summary>
[DisplayName("Insert Table Rows")]
[Description("Adds rows to a table, before a given row or after the last one.")]
public sealed class InsertTableRows : BaseNativeChild
{
    /// <summary>Which table. Tables are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table to add to. Tables are numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <summary>Row to insert before, when the option is Before.</summary>
    [Category("Input")]
    [DisplayName("Row Position")]
    [Description("Row to insert before, when Row Option is Before. Rows are numbered from 1.")]
    public InArgument<int> RowPosition { get; set; } = null!;

    /// <summary>How many rows to add.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("No Rows")]
    [Description("How many rows to add.")]
    public InArgument<int> NoRows { get; set; } = null!;

    /// <summary>Height for the new rows.</summary>
    [Category("Input")]
    [DisplayName("Row Height")]
    [Description("Height in points for the new rows. Zero leaves the default.")]
    public InArgument<float> RowHeight { get; set; } = null!;

    /// <summary>Whether to insert before a row or after the last.</summary>
    [Category("Input")]
    [DisplayName("Row Option")]
    [Description("Insert before Row Position, or after the last row.")]
    public EnumInsertOption RowOption { get; set; } = EnumInsertOption.AfterLast;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.InsertTableRows(
            TableIndex.Get(context),
            RowPosition?.Get(context) ?? 0,
            NoRows.Get(context),
            RowHeight?.Get(context) ?? 0,
            RowOption);
}
