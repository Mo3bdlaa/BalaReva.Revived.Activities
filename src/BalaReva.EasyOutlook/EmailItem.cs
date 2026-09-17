// Aliased as Interop, not Outlook: this package also declares a BalaReva.Outlook
// namespace, and inside BalaReva.EasyOutlook that sibling wins over a using alias.
using Interop = Microsoft.Office.Interop.Outlook;

namespace BalaReva.Outlook;

/// <summary>
/// A mail message returned by <c>GetMailMessages</c>.
/// </summary>
/// <remarks>
/// Note the namespace: the published package put this type in <c>BalaReva.Outlook</c>
/// rather than <c>BalaReva.EasyOutlook</c> like everything else. That looks like an
/// oversight, but a workflow binds by full type name, so it has to stay where it is.
///
/// This is a thin wrapper: the COM object is handed straight through so a workflow can
/// reach the whole Outlook object model from an expression. It is also why the package
/// cannot be made portable.
/// </remarks>
public sealed class EmailItem
{
    /// <summary>The underlying Outlook mail item.</summary>
    public Interop.MailItem? OutlookEmailItem { get; set; }
}
