using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Tables;

/// <summary>Adds columns to a table.</summary>
[DisplayName("Insert Table Columns")]
[Description("Adds columns to a table, before a given column or after the last one.")]
public sealed class InsertTableColumns : BaseNativeChild
{
    /// <summary>Which table. Tables are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table to add to. Tables are numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <summary>Column to insert before, when the option is Before.</summary>
    [Category("Input")]
    [DisplayName("Column Position")]
    [Description("Column to insert before, when Column Option is Before. Columns are numbered from 1.")]
    public InArgument<int> ColumnPosition { get; set; } = null!;

    /// <summary>How many columns to add.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("No Columns")]
    [Description("How many columns to add.")]
    public InArgument<int> NoColumns { get; set; } = null!;

    /// <summary>Width for the new columns.</summary>
    [Category("Input")]
    [DisplayName("Column Width")]
    [Description("Width in points for the new columns. Zero leaves the default.")]
    public InArgument<float> ColumnWidth { get; set; } = null!;

    /// <summary>Whether to insert before a column or after the last.</summary>
    [Category("Input")]
    [DisplayName("Column Option")]
    [Description("Insert before Column Position, or after the last column.")]
    public EnumInsertOption ColumnOption { get; set; } = EnumInsertOption.AfterLast;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.InsertTableColumns(
            TableIndex.Get(context),
            ColumnPosition?.Get(context) ?? 0,
            NoColumns.Get(context),
            ColumnWidth?.Get(context) ?? 0,
            ColumnOption);
}
