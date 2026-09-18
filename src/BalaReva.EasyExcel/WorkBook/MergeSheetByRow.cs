using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.WorkBook;

/// <summary>Appends the rows of another workbook's sheet to one of this workbook's.</summary>
[DisplayName("Merge Sheet By Row")]
[Description("Appends the rows of another workbook's sheet to one of this workbook's.")]
public sealed class MergeSheetByRow : ExcelActivity
{
    /// <summary>Range to append. Empty means everything with anything in it.</summary>
    [Category("Input")]
    [DisplayName("Append Cell Range")]
    [Description("Range to append. Empty means everything with anything in it.")]
    public InArgument<string> AppendCellRange { get; set; } = null!;

    /// <summary>Full path of the workbook to append from.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Append File Name")]
    [Description("Full path of the workbook to append from.")]
    public InArgument<string> AppendFileName { get; set; } = null!;

    /// <summary>Password needed to open it, if it has one.</summary>
    [Category("Input")]
    [DisplayName("Append File Password")]
    [Description("Password needed to open it, if it has one.")]
    public InArgument<string> AppendFilePassword { get; set; } = null!;

    /// <summary>Password needed to change it, if it has one.</summary>
    [Category("Input")]
    [DisplayName("Append Modify Password")]
    [Description("Password needed to change it, if it has one.")]
    public InArgument<string> AppendModifyPassword { get; set; } = null!;

    /// <summary>Sheet to append from. Empty means its active sheet.</summary>
    [Category("Input")]
    [DisplayName("Append Sheet")]
    [Description("Sheet to append from. Empty means its active sheet.")]
    public InArgument<string> AppendSheet { get; set; } = null!;

    /// <summary>Sheet to append to. Empty means the active sheet.</summary>
    [Category("Input")]
    [DisplayName("Sheet Name")]
    [Description("Sheet to append to. Empty means the active sheet.")]
    public InArgument<string> SheetName { get; set; } = null!;

    /// <summary>Column to start writing at, for example A.</summary>
    [Category("Input")]
    [DisplayName("Start Column")]
    [Description("Column to start writing at, for example A.")]
    public InArgument<string> StartColumn { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook)
        => workbook.MergeSheetByRow(
            SheetName?.Get(context) ?? string.Empty,
            new MergeRequest
            {
                AppendFileName = Require(context, AppendFileName, nameof(AppendFileName)),
                AppendFilePassword = AppendFilePassword?.Get(context) ?? string.Empty,
                AppendModifyPassword = AppendModifyPassword?.Get(context) ?? string.Empty,
                AppendSheet = AppendSheet?.Get(context) ?? string.Empty,
                AppendCellRange = AppendCellRange?.Get(context) ?? string.Empty,
                StartColumn = StartColumn?.Get(context) ?? string.Empty,
            });
}
