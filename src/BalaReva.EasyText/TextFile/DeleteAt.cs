using System.Activities;
using System.ComponentModel;
using BalaReva.EasyText.Base;

namespace BalaReva.EasyText.TextFile;

/// <summary>Deletes a line from the scoped file.</summary>
[DisplayName("Delete At")]
[Description("Deletes the line at the given position in the scoped text file.")]
public sealed class DeleteAt : BaseChildActivity
{
    /// <summary>1-based line number to delete.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Delete Line Index")]
    [Description("1-based number of the line to delete.")]
    public InArgument<int> DeleteLineIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteChild(CodeActivityContext context, TextDocument document)
    {
        var lines = document.ReadAllLines().ToList();
        var index = LineIndex(DeleteLineIndex.Get(context), lines.Count, nameof(DeleteLineIndex));
        lines.RemoveAt(index);
        document.WriteAllLines(lines);
    }
}
