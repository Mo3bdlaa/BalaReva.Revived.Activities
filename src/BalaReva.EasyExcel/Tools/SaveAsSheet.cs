using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Tools;

/// <summary>Saves one sheet of a workbook out as its own file.</summary>
/// <remarks>
/// Stands outside any scope: it opens the workbook it is given, writes the sheet out and
/// closes it again.
/// </remarks>
[DisplayName("Save As Sheet")]
[Description("Saves one sheet of a workbook out as its own file.")]
public sealed class SaveAsSheet : CodeActivity
{
    /// <summary>Full path of the workbook to read.</summary>
    [Category("Input")]
    [DisplayName("File Name")]
    [Description("Full path of the workbook to read.")]
    public InArgument<string> FileName { get; set; } = null!;

    /// <summary>Password needed to open it, if it has one.</summary>
    [Category("Input")]
    [DisplayName("File Password")]
    [Description("Password needed to open it, if it has one.")]
    public InArgument<string> FilePassword { get; set; } = null!;

    /// <summary>Password needed to change it, if it has one.</summary>
    [Category("Input")]
    [DisplayName("Modify Password")]
    [Description("Password needed to change it, if it has one.")]
    public InArgument<string> ModifyPassword { get; set; } = null!;

    /// <summary>Full path of the file to write.</summary>
    [Category("Input")]
    [DisplayName("New File Name")]
    [Description("Full path of the file to write.")]
    public InArgument<string> NewFileName { get; set; } = null!;

    /// <summary>Sheet to write out. Empty means the active sheet.</summary>
    [Category("Input")]
    [DisplayName("Sheet")]
    [Description("Sheet to write out. Empty means the active sheet.")]
    public InArgument<string> Sheet { get; set; } = null!;

    /// <inheritdoc />
    protected override void Execute(CodeActivityContext context)
    {
        var fileName = FileName?.Get(context);
        if (string.IsNullOrWhiteSpace(fileName))
            throw new ArgumentException("File Name is required.", nameof(FileName));

        var newFileName = NewFileName?.Get(context);
        if (string.IsNullOrWhiteSpace(newFileName))
            throw new ArgumentException("New File Name is required.", nameof(NewFileName));

        (context.GetExtension<IExcelService>() ?? ExcelService.Instance).SaveAsSheet(
            new SaveSheetRequest
            {
                FileName = fileName,
                FilePassword = FilePassword?.Get(context) ?? string.Empty,
                ModifyPassword = ModifyPassword?.Get(context) ?? string.Empty,
                NewFileName = newFileName,
                Sheet = Sheet?.Get(context) ?? string.Empty,
            });
    }
}
