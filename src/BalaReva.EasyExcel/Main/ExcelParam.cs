namespace BalaReva.EasyExcel.Main;

/// <summary>
/// The workbook an <c>ExcelScope</c> is open on, handed to its body.
/// </summary>
/// <remarks>
/// Note <c>ModiPassword</c> rather than ModifyPassword: that is how the published
/// package spelled it, and a workflow binds by property name.
/// </remarks>
public sealed class ExcelParam
{
    /// <summary>
    /// The underlying COM workbook, for workflows that need to reach past these
    /// activities and drive Excel themselves.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Published as a <c>Microsoft.Office.Interop.Excel.Workbook</c>; declared here as an
    /// object, so a workflow that wants the interop type has to cast. This is the same
    /// call made for <c>PowerPointObject.PptPersentation</c> and for the same reason:
    /// every Office interop assembly on NuGet has a hard reference on 'office'
    /// (Microsoft.Office.Core), which Microsoft publishes nowhere, and the CLR looks for
    /// it as soon as it loads a member typed that way. The scope hands this object to its
    /// body, so typing it Workbook would stop the scope loading at all on a machine
    /// without office.dll rather than only affecting this one property.
    /// </para>
    /// <para>Null unless the scope is running against the real <see cref="ExcelService"/>.</para>
    /// </remarks>
    public object? ExcelWorkBook { get; set; }

    /// <summary>Full path of the open workbook.</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>Password used to permit changes to the workbook.</summary>
    public string ModiPassword { get; set; } = string.Empty;

    /// <summary>Password used to open the workbook.</summary>
    public string Password { get; set; } = string.Empty;
}
