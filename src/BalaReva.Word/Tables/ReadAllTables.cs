using System.Activities;
using System.ComponentModel;
using System.Data;

namespace BalaReva.Word.Tables;

/// <summary>Reads every table in the document into a DataSet.</summary>
[DisplayName("Read All Tables")]
[Description("Reads every table in the document, one DataTable each.")]
public sealed class ReadAllTables : BaseNativeChild
{
    /// <summary>Whether the first row of each table holds column names.</summary>
    [Category("Input")]
    [DisplayName("With Header")]
    [Description("Treat each table's first row as its column names.")]
    public bool WithHeader { get; set; }

    /// <summary>One DataTable per Word table.</summary>
    [Category("Output")]
    [DisplayName("Result Set")]
    [Description("One DataTable per table in the document.")]
    public OutArgument<DataSet> ResultSet { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        ResultSet.Set(context, document.ReadAllTables(WithHeader));
}
