using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Settings.TrustedLocation;

/// <summary>Adds a folder to Excel's trusted locations.</summary>
[DisplayName("Add Trusted Location")]
[Description("Adds a folder to Excel's trusted locations.")]
public sealed class AddTrustedLocation : ExcelActivity
{
    /// <summary>Note to record against the entry.</summary>
    [Category("Input")]
    [DisplayName("Description")]
    [Description("Note to record against the entry.")]
    public InArgument<string> Description { get; set; } = null!;

    /// <summary>Folder to trust.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Folder Path")]
    [Description("Folder to trust.")]
    public InArgument<string> FolderPath { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook)
        => workbook.AddTrustedLocation(
            Require(context, FolderPath, nameof(FolderPath)),
            Description?.Get(context) ?? string.Empty);
}
