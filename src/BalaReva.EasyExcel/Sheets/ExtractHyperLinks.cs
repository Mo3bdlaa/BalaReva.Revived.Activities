using System.Activities;
using System.ComponentModel;
using System.Data;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Reads every hyperlink in a range.</summary>
[DisplayName("Extract Hyper Links")]
[Description("Reads every hyperlink in a range.")]
public sealed class ExtractHyperLinks : BaseActivity
{
    /// <summary>Range to read. Empty means the whole sheet.</summary>
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to read. Empty means the whole sheet.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>Where each hyperlink points.</summary>
    [Category("Output")]
    [DisplayName("Result")]
    [Description("Where each hyperlink points.")]
    public OutArgument<string[]> Result { get; set; } = null!;

    /// <summary>One row per hyperlink, with its cell, text and address.</summary>
    [Category("Output")]
    [DisplayName("Url Table")]
    [Description("One row per hyperlink, with its cell, text and address.")]
    public OutArgument<DataTable> UrlTable { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
    {
        var (addresses, table) = workbook.ExtractHyperLinks(
            sheetName, CellRange?.Get(context) ?? string.Empty);

        Result.Set(context, addresses);
        UrlTable.Set(context, table);
    }
}
