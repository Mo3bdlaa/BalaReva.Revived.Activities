using System.Data;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using BalaReva.EasyExcel.Utilities;
using Interop = Microsoft.Office.Interop.Excel;

namespace BalaReva.EasyExcel;

/// <summary>Constants Excel takes as plain numbers through late-bound members.</summary>
/// <remarks>
/// A handful of Excel's members are typed in terms of <c>Microsoft.Office.Core</c>, which
/// Microsoft publishes nowhere, so those calls go through <c>dynamic</c> and take the
/// numbers directly. The rest of this file is ordinary typed interop.
/// </remarks>
internal static class Mso
{
    /// <summary>MsoTriState.msoFalse.</summary>
    internal const int False = 0;

    /// <summary>MsoTriState.msoTrue.</summary>
    internal const int True = -1;

    /// <summary>MsoShapeType.msoLinkedPicture.</summary>
    internal const int LinkedPicture = 11;

    /// <summary>MsoShapeType.msoPicture.</summary>
    internal const int Picture = 13;

    /// <summary>MsoAutomationSecurity.msoAutomationSecurityLow.</summary>
    internal const int SecurityLow = 1;

    /// <summary>MsoAutomationSecurity.msoAutomationSecurityForceDisable.</summary>
    internal const int SecurityForceDisable = 3;

    /// <summary>The MsoTriState value for a boolean.</summary>
    internal static int Tri(bool value) => value ? True : False;
}

/// <summary>Drives Excel through COM.</summary>
/// <remarks>
/// Not covered by any automated test: no build agent has Excel installed. It is kept to a
/// mechanical translation for that reason, with the judgement pushed up into the
/// activities, which are tested against a stand-in.
/// </remarks>
public sealed class ExcelService : IExcelService
{
    /// <summary>The service the activities use when no extension is registered.</summary>
    public static IExcelService Instance { get; } = new ExcelService();

    /// <inheritdoc />
    public IExcelWorkbook Open(ExcelOpenRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        return new ExcelWorkbook(request);
    }

    /// <inheritdoc />
    public void CloseAllExcel()
    {
        foreach (var process in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
        {
            using (process)
            {
                try { process.Kill(entireProcessTree: true); }
                catch (InvalidOperationException) { /* already gone */ }
                catch (System.ComponentModel.Win32Exception) { /* not ours to kill */ }
            }
        }
    }

    /// <inheritdoc />
    public void SaveAsSheet(SaveSheetRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        using var source = new ExcelWorkbook(new ExcelOpenRequest
        {
            FilePath = request.FileName,
            FilePassword = request.FilePassword,
            ModifyPassword = request.ModifyPassword,
        });

        source.SaveSheetAs(request.Sheet, request.NewFileName);
    }
}
