using System.Activities;
using System.ComponentModel;
using BalaReva.EasyText.Base;

namespace BalaReva.EasyText.TextFile;

/// <summary>Inserts a line into the scoped file.</summary>
[DisplayName("Insert Line At")]
[Description("Inserts a line at the given position in the scoped text file.")]
public sealed class InsertLineAt : BaseChildActivity
{
    /// <summary>
    /// 1-based line number the new line takes. Existing lines from here down shift
    /// by one; passing line count + 1 appends.
    /// </summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Line At")]
    [Description("1-based position the new line takes. Pass line count + 1 to append.")]
    public InArgument<int> LineAt { get; set; } = null!;

    /// <summary>Text of the new line.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Line Text")]
    [Description("Text of the new line.")]
    public InArgument<string> LineText { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteChild(CodeActivityContext context, TextDocument document)
    {
        var lines = document.ReadAllLines().ToList();
        var at = LineAt.Get(context);
        // Inserting one past the end is an append, so the usual range check is off by one here.
        if (at < 1 || at > lines.Count + 1)
        {
            throw new ArgumentOutOfRangeException(
                nameof(LineAt), at, $"Line At must be between 1 and {lines.Count + 1}.");
        }

        lines.Insert(at - 1, LineText.Get(context) ?? string.Empty);
        document.WriteAllLines(lines);
    }
}
