using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Tables;

/// <summary>Sets the height of table rows.</summary>
[DisplayName("Row Height")]
[Description("Sets the height and height rule of table rows.")]
public sealed class RowHeight : BaseNativeChild
{
    /// <summary>Which table. Tables are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table. Tables are numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <summary>Rows to set. Empty means every row.</summary>
    [Category("Input")]
    [DisplayName("Row Indexes")]
    [Description("Rows to set, numbered from 1. Leave empty for every row.")]
    public InArgument<int[]> RowIndexes { get; set; } = null!;

    /// <summary>Height in points.</summary>
    [Category("Input")]
    [DisplayName("Row Height")]
    [Description("Height in points. Ignored when the rule is Auto.")]
    public InArgument<float> tblRowHeight { get; set; } = null!;

    /// <summary>How the height is applied.</summary>
    [Category("Input")]
    [DisplayName("Height Rule")]
    [Description("Auto sizes to content; AtLeast and Exactly use the given height.")]
    public EnumRowHeightRule HeightRule { get; set; } = EnumRowHeightRule.Auto;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.RowHeight(
            TableIndex.Get(context),
            RowIndexes?.Get(context) ?? [],
            tblRowHeight?.Get(context) ?? 0,
            HeightRule);
}
