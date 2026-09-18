namespace BalaReva.Revived.Shared;

/// <summary>
/// Resolves where an archive entry is allowed to land, and refuses anything that would
/// land outside the extraction folder.
/// </summary>
/// <remarks>
/// <para>
/// This is the whole point of these two packages being reimplemented. Every advisory
/// against their published dependencies is the same bug: an archive entry whose name is
/// <c>../../etc/passwd</c>, or an absolute path, or a Windows drive-qualified path, and a
/// library that joins it to the destination folder without checking where it ends up.
/// SharpZipLib 1.2.0 had it twice (CVE-2021-32840, CVE-2021-32842), SharpCompress 0.26.0
/// twice (CVE-2021-39208, CVE-2026-44788), and DotNetZip has had it since 2018
/// (CVE-2018-1002205) with no fix, because it has not been maintained since 2011.
/// </para>
/// <para>
/// Current versions of the maintained libraries defend themselves. This checks anyway:
/// the class of bug has recurred in every one of them at least once, the check is cheap,
/// and it means an advisory against a future version of a dependency is not automatically
/// an advisory against these activities.
/// </para>
/// </remarks>
internal static class SafeExtractPath
{
    /// <summary>
    /// The full path an entry may be written to, or null when the entry is a directory
    /// marker with nothing to write.
    /// </summary>
    /// <param name="destinationFolder">Folder the caller asked to extract into.</param>
    /// <param name="entryName">Entry name exactly as the archive gives it.</param>
    /// <exception cref="UnauthorizedAccessException">
    /// The entry would land outside <paramref name="destinationFolder"/>.
    /// </exception>
    public static string Resolve(string destinationFolder, string entryName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationFolder);

        if (string.IsNullOrWhiteSpace(entryName))
            throw new UnauthorizedAccessException("The archive holds an entry with no name.");

        // A rooted or drive-qualified name ignores the destination altogether, so
        // Path.Combine would hand back the entry's own path.
        if (Path.IsPathRooted(entryName) || HasDriveOrUnc(entryName))
            throw new UnauthorizedAccessException(
                $"The archive holds an absolute entry, '{entryName}', which would be written "
                + "outside the extraction folder. Refusing to extract it.");

        var root = Path.TrimEndingDirectorySeparator(Path.GetFullPath(destinationFolder));
        var target = Path.GetFullPath(Path.Combine(root, Normalize(entryName)));

        // The separator matters: without it, extracting into /tmp/out would accept an
        // entry resolving to /tmp/outside/evil, which shares the prefix but not the folder.
        if (target != root &&
            !target.StartsWith(root + Path.DirectorySeparatorChar, PathComparison))
        {
            throw new UnauthorizedAccessException(
                $"The archive holds an entry, '{entryName}', that would be written outside "
                + "the extraction folder. Refusing to extract it.");
        }

        return target;
    }

    /// <summary>Creates the folder an entry is going into, once its path is known safe.</summary>
    public static void EnsureFolder(string resolvedPath)
    {
        var folder = Path.GetDirectoryName(resolvedPath);
        if (!string.IsNullOrEmpty(folder)) Directory.CreateDirectory(folder);
    }

    /// <summary>Whether an entry name carries a drive letter or a UNC prefix.</summary>
    /// <remarks>
    /// Path.IsPathRooted does not catch "C:evil.txt" on Linux, and an archive built on
    /// Windows can carry either. Checking by hand keeps the answer the same on both.
    /// </remarks>
    private static bool HasDriveOrUnc(string entryName) =>
        (entryName.Length >= 2 && entryName[1] == ':' && char.IsLetter(entryName[0]))
        || entryName.StartsWith("\\\\", StringComparison.Ordinal)
        || entryName.StartsWith("//", StringComparison.Ordinal);

    /// <summary>Archives always use forward slashes; the local platform may not.</summary>
    private static string Normalize(string entryName) =>
        entryName.Replace('/', Path.DirectorySeparatorChar)
                 .Replace('\\', Path.DirectorySeparatorChar);

    /// <summary>
    /// Case-insensitive off Linux, matching how each platform decides two paths are the
    /// same file.
    /// </summary>
    private static StringComparison PathComparison =>
        OperatingSystem.IsLinux() ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
}
