using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Tables;

/// <summary>Applies an autofit behaviour to a table.</summary>
[DisplayName("Table Auto Fit")]
[Description("Applies an autofit behaviour to a table.")]
public sealed class TableAutoFit : BaseNativeChild
{
    /// <summary>Which table. Tables are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table. Tables are numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <summary>How the table sizes itself.</summary>
    [Category("Input")]
    [DisplayName("Auto Fit")]
    [Description("Fit to a fixed column width, to the content, or to the window.")]
    public EnumAutoFitBehavior AutoFit { get; set; } = EnumAutoFitBehavior.FitContent;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.TableAutoFit(TableIndex.Get(context), AutoFit);
}
