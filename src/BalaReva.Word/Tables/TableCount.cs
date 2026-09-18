using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Tables;

/// <summary>Counts the tables in the document.</summary>
[DisplayName("Table Count")]
[Description("Counts the tables in the document.")]
public sealed class TableCount : BaseNativeChild
{
    /// <summary>Number of tables.</summary>
    [Category("Output")]
    [DisplayName("Result")]
    [Description("Number of tables in the document.")]
    public OutArgument<int> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        Result.Set(context, document.TableCount());
}
