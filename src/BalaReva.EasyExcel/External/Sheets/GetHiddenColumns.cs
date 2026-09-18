using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.External.Sheets;

/// <summary>Lists the hidden columns of a sheet.</summary>
/// <remarks>
/// Reads the file directly rather than through Excel, so it needs no Excel installation
/// and no scope.
/// </remarks>
[DisplayName("Get Hidden Columns")]
[Description("Lists the hidden columns of a sheet.")]
public sealed class GetHiddenColumns : BaseOpenXml
{
    /// <summary>The hidden columns, by letter.</summary>
    [RequiredArgument]
    [Category("Output")]
    [DisplayName("Result")]
    [Description("The hidden columns, by letter.")]
    public OutArgument<List<string>> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IOpenXmlReader reader, string filePath, string sheetName)
        => Result.Set(context, reader.HiddenColumns(filePath, sheetName));
}
