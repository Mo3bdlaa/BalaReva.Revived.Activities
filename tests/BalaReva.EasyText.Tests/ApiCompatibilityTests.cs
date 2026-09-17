using System.Reflection;
using System.Text.Json;
using BalaReva.EasyText.Base;

namespace BalaReva.EasyText.Tests;

/// <summary>
/// Checks the reimplementation against the recorded surface of the published
/// BalaReva.EasyText.Activities 3.0.1.
/// </summary>
/// <remarks>
/// An existing .xaml binds to an activity by type full name and to its arguments by
/// property name, so a rename here silently breaks every workflow that upgrades.
/// This is the test that makes "drop-in replacement" a claim rather than a hope.
///
/// Design-time types are excluded on purpose: the published package shipped WPF
/// designers, this one does not, and Studio falls back to a generated property grid.
/// That changes how the activity looks, not what a workflow binds to.
/// </remarks>
public class ApiCompatibilityTests
{
    private const string PackageId = "BalaReva.EasyText.Activities";

    private static readonly Assembly Revived = typeof(BaseChildActivity).Assembly;

    public static TheoryData<string> OriginalActivities()
    {
        var data = new TheoryData<string>();
        foreach (var type in Original()) data.Add(type.GetProperty("FullName")!.GetString()!);
        return data;
    }

    [Theory]
    [MemberData(nameof(OriginalActivities))]
    public void Every_published_activity_still_exists_under_the_same_full_name(string fullName) =>
        Assert.True(Revived.GetType(fullName) is not null,
            $"{fullName} existed in the published package but not here; workflows binding to it would break.");

    [Theory]
    [MemberData(nameof(OriginalActivities))]
    public void Every_published_property_still_exists_with_a_compatible_type(string fullName)
    {
        var original = Original().Single(t => t.GetProperty("FullName")!.GetString() == fullName);
        var revived = Revived.GetType(fullName);
        Assert.NotNull(revived);

        foreach (var property in original.GetProperty("Properties")!.EnumerateArray())
        {
            var name = property.GetProperty("Name").GetString()!;
            var expected = Normalize(property.GetProperty("Type").GetString()!);

            var actual = revived!.GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
            Assert.True(actual is not null, $"{fullName}.{name} is missing.");
            Assert.Equal(expected, Describe(actual!.PropertyType));
        }
    }

    [Fact]
    public void The_recorded_surface_was_actually_loaded()
    {
        // Guards against this whole suite silently passing on an empty list if the
        // audit artefact ever stops being copied next to the tests.
        Assert.Equal(13, Original().Count);
    }

    private static List<JsonElement> Original()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "api-surface.json");
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        return document.RootElement
            .GetProperty("packages")
            .EnumerateArray()
            .Single(p => p.GetProperty("Id").GetString() == PackageId)
            .GetProperty("Activities")
            .EnumerateArray()
            .Where(t => !t.GetProperty("Namespace").GetString()!.Contains(".Design", StringComparison.Ordinal))
            .Select(t => t.Clone())
            .ToList();
    }

    /// <summary>Renders a recorded signature string as "Outer&lt;Inner&gt;".</summary>
    private static string Normalize(string recorded)
    {
        var open = recorded.IndexOf('<');
        if (open < 0) return Leaf(recorded);
        var outer = Leaf(recorded[..open]);
        var inner = recorded[(open + 1)..recorded.LastIndexOf('>')];
        return $"{outer}<{Leaf(inner)}>";
    }

    /// <summary>Renders a runtime type the same way <see cref="Normalize"/> renders a recorded one.</summary>
    private static string Describe(Type type)
    {
        if (!type.IsGenericType) return Leaf(type.Name);
        var outer = type.Name[..type.Name.IndexOf('`')];
        return $"{outer}<{Leaf(type.GetGenericArguments()[0].Name)}>";
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
        name = name switch
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
        };
        return name + suffix;
    }
}
