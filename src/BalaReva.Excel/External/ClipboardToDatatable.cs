using System.Activities;
using System.ComponentModel;
using System.Data;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.External;

/// <summary>Reads tabular data off the clipboard.</summary>
/// <remarks>
/// The only activity in the package that opens no workbook, so it carries no file path,
/// sheet name or ExecutionResult.
/// </remarks>
[DisplayName("Clipboard To Datatable")]
[Description("Reads whatever tabular data is on the clipboard into a DataTable.")]
public sealed class ClipboardToDatatable : BaseExcel
{
    /// <summary>Whether the first row holds column names.</summary>
    [Category("Input")]
    [DisplayName("Has Header")]
    [Description("Treat the first row as column names.")]
    public InArgument<bool> HasHeader { get; set; } = null!;

    /// <summary>The clipboard contents.</summary>
    [Category("Output")]
    [DisplayName("Datatable")]
    [Description("The clipboard contents as a DataTable.")]
    public OutArgument<DataTable> Datatable { get; set; } = null!;

    /// <inheritdoc />
    protected override void Execute(CodeActivityContext context)
    {
        var service = context.GetExtension<IExcelService>() ?? ExcelService.Instance;
        Datatable.Set(context, service.ClipboardToDataTable(HasHeader?.Get(context) ?? false));
    }
}
