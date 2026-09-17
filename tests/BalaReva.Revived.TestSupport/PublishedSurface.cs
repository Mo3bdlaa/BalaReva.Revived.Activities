using System.Reflection;
using System.Text.Json;
using Xunit;

namespace BalaReva.Revived.TestSupport;

/// <summary>
/// Checks a revived assembly against the recorded surface of the package it replaces.
/// </summary>
/// <remarks>
/// A .xaml binds to an activity by type full name and to each input and output by
/// property name, so a rename silently breaks every workflow that upgrades. This is
/// what makes "drop-in replacement" a claim rather than a hope.
///
/// Design-time types are excluded: the published packages shipped WPF designers, these
/// do not, and Studio falls back to a generated property grid. That changes how an
/// activity looks, not what a workflow binds to.
/// </remarks>
public static class PublishedSurface
{
    private static readonly Lazy<JsonDocument> Recorded = new(() =>
        JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "api-surface.json"))));

    /// <summary>The activity types the published <paramref name="packageId"/> exposed.</summary>
    public static List<JsonElement> Activities(string packageId) =>
        Recorded.Value.RootElement
            .GetProperty("packages")
            .EnumerateArray()
            .Single(p => p.GetProperty("Id").GetString() == packageId)
            .GetProperty("Activities")
            .EnumerateArray()
            .Where(t => !t.GetProperty("Namespace").GetString()!.Contains(".Design", StringComparison.Ordinal))
            .Select(t => t.Clone())
            .ToList();

    /// <summary>Full names of those activity types, for use as xUnit theory data.</summary>
    public static TheoryData<string> ActivityNames(string packageId)
    {
        var data = new TheoryData<string>();
        foreach (var type in Activities(packageId)) data.Add(type.GetProperty("FullName")!.GetString()!);
        return data;
    }

    /// <summary>Asserts <paramref name="assembly"/> still declares <paramref name="fullName"/>.</summary>
    public static Type RequireType(Assembly assembly, string fullName)
    {
        var type = assembly.GetType(fullName);
        Assert.True(type is not null,
            $"{fullName} existed in the published package but not here; " +
            "workflows binding to it would break.");
        return type!;
    }

    /// <summary>
    /// Asserts every property the published type exposed is still present, with a
    /// type that renders the same.
    /// </summary>
    public static void RequireProperties(Assembly assembly, string packageId, string fullName)
    {
        var original = Activities(packageId).Single(t => t.GetProperty("FullName").GetString() == fullName);
        var revived = RequireType(assembly, fullName);

        foreach (var property in original.GetProperty("Properties").EnumerateArray())
        {
            var name = property.GetProperty("Name").GetString()!;
            var expected = Normalize(property.GetProperty("Type").GetString()!);

            var actual = revived.GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
            Assert.True(actual is not null, $"{fullName}.{name} is missing.");
            Assert.Equal(expected, Describe(actual!.PropertyType));
        }
    }

    /// <summary>
    /// Renders a recorded signature string as "Outer&lt;Inner&gt;", recursing so that
    /// a nested generic such as InArgument&lt;List&lt;string&gt;&gt; renders in full.
    /// </summary>
    private static string Normalize(string recorded)
    {
        var open = recorded.IndexOf('<');
        if (open < 0) return Leaf(recorded);
        var outer = Leaf(recorded[..open]);
        var inner = recorded[(open + 1)..recorded.LastIndexOf('>')];
        return $"{outer}<{string.Join(", ", SplitArguments(inner).Select(Normalize))}>";
    }

    /// <summary>Renders a runtime type the same way <see cref="Normalize"/> renders a recorded one.</summary>
    private static string Describe(Type type)
    {
        if (type.IsArray) return Describe(type.GetElementType()!) + "[]";
        if (!type.IsGenericType) return Leaf(type.Name);
        var outer = type.Name[..type.Name.IndexOf('`')];
        return $"{outer}<{string.Join(", ", type.GetGenericArguments().Select(Describe))}>";
    }

    /// <summary>Splits generic arguments on commas that are not inside a nested pair.</summary>
    private static List<string> SplitArguments(string inner)
    {
        var parts = new List<string>();
        var depth = 0;
        var start = 0;
        for (var i = 0; i < inner.Length; i++)
        {
            if (inner[i] == '<') depth++;
            else if (inner[i] == '>') depth--;
            else if (inner[i] == ',' && depth == 0)
            {
                parts.Add(inner[start..i].Trim());
                start = i + 1;
            }
        }
        parts.Add(inner[start..].Trim());
        return parts;
    }

    /// <summary>Strips namespaces and maps C# keywords onto their CLR type names.</summary>
    private static string Leaf(string name)
    {
        var suffix = string.Empty;
        while (name.EndsWith("[]", StringComparison.Ordinal))
        {
            suffix += "[]";
            name = name[..^2];
        }
        name = name[(name.LastIndexOf('.') + 1)..];
        var arity = name.IndexOf('`');
        if (arity >= 0) name = name[..arity];
        return (name switch
        {
            "bool" => "Boolean",
            "int" => "Int32",
            "long" => "Int64",
            "short" => "Int16",
            "byte" => "Byte",
            "double" => "Double",
            "float" => "Single",
            "string" => "String",
            "object" => "Object",
            _ => name,
        }) + suffix;
    }
}
