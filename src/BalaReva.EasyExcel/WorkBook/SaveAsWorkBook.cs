using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.WorkBook;

/// <summary>Saves a copy of the workbook under another name.</summary>
[DisplayName("Save As WorkBook")]
[Description("Saves a copy of the workbook under another name.")]
public sealed class SaveAsWorkBook : ExcelActivity
{
    /// <summary>Format to write it in.</summary>
    [Category("Input")]
    [DisplayName("File Format")]
    [Description("Format to write it in.")]
    public FileFormatEnum FileFormat { get; set; } = FileFormatEnum.CurrentPlatformText;

    /// <summary>Full path of the file to write.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("File Name")]
    [Description("Full path of the file to write.")]
    public InArgument<string> FileName { get; set; } = null!;

    /// <summary>Password to require for opening the copy.</summary>
    [Category("Input")]
    [DisplayName("Password")]
    [Description("Password to require for opening the copy.")]
    public InArgument<string> Password { get; set; } = null!;

    /// <summary>Password to require for changing the copy.</summary>
    [Category("Input")]
    [DisplayName("Write Res Password")]
    [Description("Password to require for changing the copy.")]
    public InArgument<string> WriteResPassword { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook)
        => workbook.SaveAsWorkBook(
            Require(context, FileName, nameof(FileName)),
            FileFormat,
            Password?.Get(context) ?? string.Empty,
            WriteResPassword?.Get(context) ?? string.Empty);
}
