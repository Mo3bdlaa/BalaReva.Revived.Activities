using System.Activities;
using System.ComponentModel;
using BalaReva.EasyText.Base;

namespace BalaReva.EasyText.TextFile;

/// <summary>Removes blank lines from the scoped file.</summary>
[DisplayName("Remove Empty Line")]
[Description("Removes every blank or whitespace-only line from the scoped text file.")]
public sealed class RemoveEmptyLine : BaseChildActivity
{
    /// <inheritdoc />
    protected override void ExecuteChild(CodeActivityContext context, TextDocument document)
    {
        var kept = document.ReadAllLines().Where(line => !string.IsNullOrWhiteSpace(line));
        document.WriteAllLines(kept);
    }
}
