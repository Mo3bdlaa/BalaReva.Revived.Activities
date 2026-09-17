using System.Activities;

namespace BalaReva.EasyText.Base;

/// <summary>
/// Resolves the <c>TextScope</c> a child activity is running inside, and produces a
/// readable error when it is not inside one at all.
/// </summary>
/// <remarks>
/// Descendants resolve the scope by handle type, so nesting scopes is allowed and
/// the innermost one wins.
/// </remarks>
public class ParentActivityValidation
{
    /// <summary>
    /// Returns the enclosing scope's open document, or throws when the activity is
    /// not inside a <c>TextScope</c>.
    /// </summary>
    public static TextDocument Resolve(CodeActivityContext context)
    {
        ArgumentNullException.ThrowIfNull(context);
        return context.GetProperty<TextScopeHandle>()?.Document
               ?? throw new InvalidOperationException(
                   "This activity must be placed inside a Text Scope activity, "
                   + "which supplies the file path.");
    }
}
