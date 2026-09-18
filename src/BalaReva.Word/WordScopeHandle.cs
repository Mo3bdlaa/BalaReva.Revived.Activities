using System.Activities;

namespace BalaReva.Word;

/// <summary>
/// The execution property a <c>WordScope</c> publishes so its children can find the
/// open document.
/// </summary>
/// <remarks>
/// Derives from <see cref="Handle"/> because that is the constraint on
/// <see cref="CodeActivityContext.GetProperty{THandle}"/>, which is how a
/// <see cref="BaseNativeChild"/> resolves its scope.
/// </remarks>
public sealed class WordScopeHandle : Handle
{
    /// <summary>The document the enclosing scope has open.</summary>
    public IWordDocument? Document { get; set; }
}
