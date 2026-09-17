using System.Activities;

namespace BalaReva.EasyText.Base;

/// <summary>
/// The workflow execution property a <c>TextScope</c> publishes so that its
/// descendants can find the file it is open on.
/// </summary>
/// <remarks>
/// This has to derive from <see cref="Handle"/> because that is the constraint on
/// <see cref="CodeActivityContext.GetProperty{THandle}"/>, which is how a
/// <see cref="BaseChildActivity"/> resolves its scope. Keeping the file logic in
/// <see cref="TextDocument"/> rather than here leaves that logic testable without
/// standing up a workflow.
/// </remarks>
public sealed class TextScopeHandle : Handle
{
    /// <summary>The file the enclosing scope is open on.</summary>
    public TextDocument? Document { get; set; }
}
