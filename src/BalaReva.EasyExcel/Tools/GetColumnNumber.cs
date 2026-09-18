using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Tools;

/// <summary>Converts Excel's column letters to a column number.</summary>
/// <remarks>
/// Arithmetic, not automation: this needs no Excel and no scope.
/// </remarks>
[DisplayName("Get Column Number")]
[Description("Converts Excel's column letters to a column number.")]
public sealed class GetColumnNumber : CodeActivity
{
    /// <summary>The number for that column, counted from 1.</summary>
    [Category("Output")]
    [DisplayName("Column Index")]
    [Description("The number for that column, counted from 1.")]
    public OutArgument<int> ColumnIndex { get; set; } = null!;

    /// <summary>Column letters, for example AA.</summary>
    [Category("Input")]
    [DisplayName("Column Name")]
    [Description("Column letters, for example AA.")]
    public InArgument<string> ColumnName { get; set; } = null!;

    /// <inheritdoc />
    protected override void Execute(CodeActivityContext context)
    {
        var name = ColumnName.Get(context);
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Column Name is required.", nameof(ColumnName));

        var index = 0;
        foreach (var letter in name.Trim().ToUpperInvariant())
        {
            if (letter is < 'A' or > 'Z')
                throw new ArgumentException(
                    $"'{name}' is not a column name.", nameof(ColumnName));
            index = (index * 26) + (letter - 'A' + 1);
        }
        ColumnIndex.Set(context, index);
    }
}
