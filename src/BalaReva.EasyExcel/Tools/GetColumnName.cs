using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Tools;

/// <summary>Converts a column number to Excel's column letters.</summary>
/// <remarks>
/// Arithmetic, not automation: this needs no Excel and no scope.
/// </remarks>
[DisplayName("Get Column Name")]
[Description("Converts a column number to Excel's column letters.")]
public sealed class GetColumnName : CodeActivity
{
    /// <summary>Column number, counted from 1.</summary>
    [Category("Input")]
    [DisplayName("Column Index")]
    [Description("Column number, counted from 1.")]
    public InArgument<int> ColumnIndex { get; set; } = null!;

    /// <summary>The letters for that column, for example AA for 27.</summary>
    [Category("Output")]
    [DisplayName("Column Name")]
    [Description("The letters for that column, for example AA for 27.")]
    public OutArgument<string> ColumnName { get; set; } = null!;

    /// <inheritdoc />
    protected override void Execute(CodeActivityContext context)
    {
        var index = ColumnIndex.Get(context);
        if (index < 1)
            throw new ArgumentOutOfRangeException(
                nameof(ColumnIndex), index, "Columns are numbered from 1.");

        var name = string.Empty;
        while (index > 0)
        {
            // Excel's columns are bijective base-26: there is no zero digit, so A
            // is 1 and Z is 26 rather than A being 0.
            var digit = (index - 1) % 26;
            name = (char)('A' + digit) + name;
            index = (index - 1) / 26;
        }
        ColumnName.Set(context, name);
    }
}
