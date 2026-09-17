namespace BalaReva.Printer.Enums;

/// <summary>
/// Access level to open a print queue with.
/// </summary>
/// <remarks>
/// The names and values match <c>System.Printing.PrintSystemDesiredAccess</c>, which
/// is what the published package used, and they must not be renumbered: a workflow can
/// persist the underlying value.
/// </remarks>
public enum AccessRightsEnum
{
    /// <summary>Not specified; the activity picks the access its operation needs.</summary>
    None = 0,
    /// <summary>Full administrative access to the print server.</summary>
    AdministrateServer = 983041,
    /// <summary>Enumerate the queues on a print server.</summary>
    EnumerateServer = 131074,
    /// <summary>Submit jobs to a printer.</summary>
    UsePrinter = 131080,
    /// <summary>Full administrative access to a printer, including pausing its queue.</summary>
    AdministratePrinter = 983052,
}
