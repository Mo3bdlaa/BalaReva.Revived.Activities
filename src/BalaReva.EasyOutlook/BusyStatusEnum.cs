namespace BalaReva.EasyOutlook;

/// <summary>How an appointment or meeting shows on the calendar.</summary>
/// <remarks>
/// Values match <c>Microsoft.Office.Interop.Outlook.OlBusyStatus</c>, which is what the
/// published package used. They must not be renumbered: a workflow can persist the
/// underlying value.
/// </remarks>
public enum BusyStatusEnum
{
    /// <summary>Shows as free.</summary>
    Free = 0,
    /// <summary>Shows as tentative.</summary>
    Tentative = 1,
    /// <summary>Shows as busy.</summary>
    Busy = 2,
    /// <summary>Shows as out of office.</summary>
    OutOfOffice = 3,
    /// <summary>Shows as working elsewhere.</summary>
    WorkingElsewhere = 4,
}
