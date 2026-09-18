using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.External.Sheets;

/// <summary>Lists the hidden rows of a sheet.</summary>
/// <remarks>
/// Reads the file directly rather than through Excel, so it needs no Excel installation
/// and no scope.
/// </remarks>
[DisplayName("Get Hidden Rows")]
[Description("Lists the hidden rows of a sheet.")]
public sealed class GetHiddenRows : BaseOpenXml
{
    /// <summary>The hidden rows, numbered from 1.</summary>
    [RequiredArgument]
    [Category("Output")]
    [DisplayName("Result")]
    [Description("The hidden rows, numbered from 1.")]
    public OutArgument<List<int>> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IOpenXmlReader reader, string filePath, string sheetName)
        => Result.Set(context, reader.HiddenRows(filePath, sheetName));
}
