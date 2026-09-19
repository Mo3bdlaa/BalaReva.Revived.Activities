using System.Collections.Immutable;
using System.IO.Compression;
using System.Reflection;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
using System.Text.Json;

namespace BalaReva.Revived.Tools.ApiSurface;

/// <summary>
/// Dumps the public API surface of the published BalaReva packages: the activity
/// types a workflow can drop onto a canvas, and the properties it binds to.
///
/// This reads metadata tables only — type, property and attribute names. It never
/// touches method bodies, because the point is to reimplement these activities
/// from scratch while keeping existing .xaml workflows binding to the same names.
/// A property renamed here is a workflow broken on upgrade.
/// </summary>
public static class Program
{
    public static int Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine("usage: ApiSurface <packages-dir> <output.json> [id-substring ...]");
            return 2;
        }

        var dir = args[0];
        var output = args[1];
        var filters = args.Skip(2).ToArray();

        var packages = Directory.EnumerateFiles(dir, "*.nupkg")
            .Where(p => filters.Length == 0 ||
                        filters.Any(f => Path.GetFileName(p)
                            .StartsWith(f + ".", StringComparison.OrdinalIgnoreCase)))
            .OrderBy(p => p)
            .Select(Describe)
            .ToList();

        var json = JsonSerializer.Serialize(
            new { packages },
            new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(output, json);

        foreach (var package in packages)
        {
            Console.WriteLine($"{package.Id,-44}{package.Activities.Count,4} activities" +
                              $"{package.Types.Count,6} public types");
        }
        Console.Error.WriteLine($"\nWrote {output}");
        return 0;
    }

    private static PackageSurface Describe(string path)
    {
        var name = Path.GetFileNameWithoutExtension(path);
        var types = new List<TypeSurface>();

        using var archive = ZipFile.OpenRead(path);
        // Prefer the modern assets; the net461 copy of the same assembly would only
        // duplicate every type.
        var entries = archive.Entries
            .Where(e => e.FullName.EndsWith(".dll", StringComparison.OrdinalIgnoreCase))
            .Where(e => !IsVendored(Path.GetFileName(e.FullName)))
            .ToList();
        var modern = entries.Where(e => !e.FullName.Contains("/net4", StringComparison.OrdinalIgnoreCase)).ToList();
        foreach (var entry in modern.Count > 0 ? modern : entries)
        {
            using var stream = entry.Open();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            buffer.Position = 0;
            try
            {
                types.AddRange(ReadTypes(buffer, entry.FullName));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"  ! {name} {entry.FullName}: {ex.Message}");
            }
        }

        return new PackageSurface(
            StripVersion(name),
            types.OrderBy(t => t.FullName).ToList(),
            types.Where(t => t.IsActivity).OrderBy(t => t.FullName).ToList());
    }

    /// <summary>
    /// Whether an assembly inside the package is a third-party one the author bundled
    /// rather than wrote.
    /// </summary>
    /// <remarks>
    /// This used to be an allow list of names starting with "BalaReva", which was wrong
    /// in a way that failed silently: BalaReva.DataTable.Activities ships its code as
    /// DataTableExtensions.Activities.dll, so the package was reported as holding zero
    /// activities rather than as unreadable. Skipping the vendored libraries by name is
    /// the same intent without the trap - an unrecognised assembly is now read, not
    /// dropped.
    /// </remarks>
    private static bool IsVendored(string fileName) =>
        VendoredPrefixes.Any(p => fileName.StartsWith(p, StringComparison.OrdinalIgnoreCase));

    private static readonly string[] VendoredPrefixes =
    [
        "Microsoft.Office.", "Interop.Microsoft.", "Office.",
        "ICSharpCode.", "SharpCompress.", "DotNetZip.", "Ionic.",
        "Npgsql.", "Newtonsoft.", "DocumentFormat.",
    ];

    private static string StripVersion(string fileName)
    {
        var i = fileName.IndexOf(".Activities.", StringComparison.Ordinal);
        return i < 0 ? fileName : fileName[..(i + ".Activities".Length)];
    }

    private static List<TypeSurface> ReadTypes(Stream stream, string assemblyPath)
    {
        var result = new List<TypeSurface>();
        using var pe = new PEReader(stream);
        if (!pe.HasMetadata) return result;
        var md = pe.GetMetadataReader();

        // Pass one: every type's immediate base, so pass two can walk the chain.
        var baseOf = new Dictionary<string, string?>(StringComparer.Ordinal);
        foreach (var handle in md.TypeDefinitions)
        {
            var t = md.GetTypeDefinition(handle);
            var ns = md.GetString(t.Namespace);
            var n = md.GetString(t.Name);
            baseOf[string.IsNullOrEmpty(ns) ? n : $"{ns}.{n}"] = ResolveTypeName(md, t.BaseType);
        }

        foreach (var handle in md.TypeDefinitions)
        {
            var type = md.GetTypeDefinition(handle);
            var attrs = type.Attributes;
            var visibility = attrs & TypeAttributes.VisibilityMask;
            if (visibility != TypeAttributes.Public && visibility != TypeAttributes.NestedPublic)
                continue;

            var ns = md.GetString(type.Namespace);
            var name = md.GetString(type.Name);
            if (name.StartsWith('<')) continue; // compiler-generated

            var baseType = ResolveTypeName(md, type.BaseType);

            // Enum members are part of the binding contract: a workflow persists the
            // member name, so renaming or reordering one breaks it on upgrade.
            var enumMembers = new List<EnumMember>();
            if (baseType == "System.Enum")
            {
                foreach (var fieldHandle in type.GetFields())
                {
                    var field = md.GetFieldDefinition(fieldHandle);
                    if ((field.Attributes & FieldAttributes.Static) == 0) continue;
                    var constHandle = field.GetDefaultValue();
                    object? value = null;
                    if (!constHandle.IsNil)
                    {
                        var constant = md.GetConstant(constHandle);
                        var reader = md.GetBlobReader(constant.Value);
                        value = constant.TypeCode switch
                        {
                            ConstantTypeCode.Int32 => reader.ReadInt32(),
                            ConstantTypeCode.Int16 => reader.ReadInt16(),
                            ConstantTypeCode.Byte => reader.ReadByte(),
                            ConstantTypeCode.SByte => reader.ReadSByte(),
                            ConstantTypeCode.Int64 => reader.ReadInt64(),
                            _ => null,
                        };
                    }
                    enumMembers.Add(new EnumMember(md.GetString(field.Name), value));
                }
            }

            var properties = new List<PropertySurface>();
            foreach (var propHandle in type.GetProperties())
            {
                var prop = md.GetPropertyDefinition(propHandle);
                var accessors = prop.GetAccessors();
                if (!IsPublic(md, accessors.Getter) && !IsPublic(md, accessors.Setter)) continue;
                properties.Add(new PropertySurface(
                    md.GetString(prop.Name),
                    DecodeSignature(md, prop),
                    accessors.Getter.IsNil ? null : IsPublic(md, accessors.Getter),
                    accessors.Setter.IsNil ? null : IsPublic(md, accessors.Setter),
                    AttributeNames(md, prop.GetCustomAttributes())));
            }

            result.Add(new TypeSurface(
                assemblyPath,
                string.IsNullOrEmpty(ns) ? name : $"{ns}.{name}",
                ns,
                name,
                baseType,
                (attrs & TypeAttributes.Abstract) != 0,
                DerivesFromActivity(baseType, baseOf),
                properties.OrderBy(p => p.Name).ToList(),
                AttributeNames(md, type.GetCustomAttributes()),
                enumMembers));
        }

        return result;
    }

    /// <summary>
    /// Walks up the inheritance chain looking for a System.Activities root. Bases that
    /// live in this same assembly are followed; anything else is matched by name, since
    /// the referenced assembly is not loaded.
    /// </summary>
    private static bool DerivesFromActivity(string? baseType, Dictionary<string, string?> baseOf)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var current = baseType;
        while (current is not null && seen.Add(current))
        {
            if (IsActivityRoot(current)) return true;
            if (!baseOf.TryGetValue(StripGenerics(current), out current)) return false;
        }
        return false;
    }

    private static string StripGenerics(string name)
    {
        var i = name.IndexOf('<');
        return i < 0 ? name : name[..i];
    }

    private static bool IsActivityRoot(string name)
    {
        var bare = StripGenerics(name);
        var leaf = bare[(bare.LastIndexOf('.') + 1)..];
        return leaf is "CodeActivity" or "AsyncCodeActivity" or "NativeActivity"
            or "Activity" or "ActivityWithResult" or "ContinuableAsyncCodeActivity";
    }

    private static bool IsPublic(MetadataReader md, MethodDefinitionHandle handle) =>
        !handle.IsNil &&
        (md.GetMethodDefinition(handle).Attributes & MethodAttributes.MemberAccessMask) == MethodAttributes.Public;

    private static List<string> AttributeNames(MetadataReader md, CustomAttributeHandleCollection handles)
    {
        var names = new List<string>();
        foreach (var handle in handles)
        {
            var attr = md.GetCustomAttribute(handle);
            string? owner = attr.Constructor.Kind switch
            {
                HandleKind.MemberReference =>
                    ResolveTypeName(md, md.GetMemberReference((MemberReferenceHandle)attr.Constructor).Parent),
                HandleKind.MethodDefinition =>
                    ResolveTypeName(md, md.GetMethodDefinition((MethodDefinitionHandle)attr.Constructor)
                        .GetDeclaringType()),
                _ => null,
            };
            if (owner is not null) names.Add(owner);
        }
        names.Sort();
        return names;
    }

    private static string? ResolveTypeName(MetadataReader md, EntityHandle handle)
    {
        if (handle.IsNil) return null;
        switch (handle.Kind)
        {
            case HandleKind.TypeDefinition:
            {
                var t = md.GetTypeDefinition((TypeDefinitionHandle)handle);
                var ns = md.GetString(t.Namespace);
                var n = md.GetString(t.Name);
                return string.IsNullOrEmpty(ns) ? n : $"{ns}.{n}";
            }
            case HandleKind.TypeReference:
            {
                var t = md.GetTypeReference((TypeReferenceHandle)handle);
                var ns = md.GetString(t.Namespace);
                var n = md.GetString(t.Name);
                return string.IsNullOrEmpty(ns) ? n : $"{ns}.{n}";
            }
            case HandleKind.TypeSpecification:
            {
                var t = md.GetTypeSpecification((TypeSpecificationHandle)handle);
                try { return t.DecodeSignature(SignatureNames.Instance, null!); }
                catch { return "<typespec>"; }
            }
            default:
                return null;
        }
    }

    private static string DecodeSignature(MetadataReader md, PropertyDefinition prop)
    {
        try { return prop.DecodeSignature(SignatureNames.Instance, null!).ReturnType; }
        catch { return "<unknown>"; }
    }
}

/// <summary>Renders metadata signatures as readable C#-ish type names.</summary>
internal sealed class SignatureNames : ISignatureTypeProvider<string, object?>
{
    public static readonly SignatureNames Instance = new();

    public string GetPrimitiveType(PrimitiveTypeCode code) => code switch
    {
        PrimitiveTypeCode.Boolean => "bool",
        PrimitiveTypeCode.Byte => "byte",
        PrimitiveTypeCode.Char => "char",
        PrimitiveTypeCode.Double => "double",
        PrimitiveTypeCode.Int16 => "short",
        PrimitiveTypeCode.Int32 => "int",
        PrimitiveTypeCode.Int64 => "long",
        PrimitiveTypeCode.Object => "object",
        PrimitiveTypeCode.SByte => "sbyte",
        PrimitiveTypeCode.Single => "float",
        PrimitiveTypeCode.String => "string",
        PrimitiveTypeCode.UInt16 => "ushort",
        PrimitiveTypeCode.UInt32 => "uint",
        PrimitiveTypeCode.UInt64 => "ulong",
        PrimitiveTypeCode.Void => "void",
        _ => code.ToString(),
    };

    private static string Name(string ns, string name) => string.IsNullOrEmpty(ns) ? name : $"{ns}.{name}";

    public string GetTypeFromDefinition(MetadataReader md, TypeDefinitionHandle h, byte rawTypeKind)
    {
        var t = md.GetTypeDefinition(h);
        return Name(md.GetString(t.Namespace), md.GetString(t.Name));
    }

    public string GetTypeFromReference(MetadataReader md, TypeReferenceHandle h, byte rawTypeKind)
    {
        var t = md.GetTypeReference(h);
        return Name(md.GetString(t.Namespace), md.GetString(t.Name));
    }

    public string GetTypeFromSpecification(MetadataReader md, object? _, TypeSpecificationHandle h, byte rawTypeKind)
        => md.GetTypeSpecification(h).DecodeSignature(this, null!);

    public string GetSZArrayType(string elementType) => elementType + "[]";
    public string GetArrayType(string elementType, ArrayShape shape) =>
        elementType + "[" + new string(',', Math.Max(shape.Rank - 1, 0)) + "]";
    public string GetPointerType(string elementType) => elementType + "*";
    public string GetByReferenceType(string elementType) => "ref " + elementType;
    public string GetPinnedType(string elementType) => elementType;
    public string GetGenericInstantiation(string genericType, ImmutableArray<string> args) =>
        Trim(genericType) + "<" + string.Join(", ", args) + ">";
    public string GetGenericMethodParameter(object? _, int index) => "!!" + index;
    public string GetGenericTypeParameter(object? _, int index) => "!" + index;
    public string GetModifiedType(string modifier, string unmodifiedType, bool isRequired) => unmodifiedType;
    public string GetFunctionPointerType(MethodSignature<string> signature) => "delegate*";

    private static string Trim(string name)
    {
        var tick = name.IndexOf('`');
        return tick < 0 ? name : name[..tick];
    }
}

internal record PropertySurface(
    string Name, string Type, bool? PublicGetter, bool? PublicSetter, List<string> Attributes);

internal record EnumMember(string Name, object? Value);

internal record TypeSurface(
    string Assembly, string FullName, string Namespace, string Name, string? BaseType,
    bool IsAbstract, bool IsActivity, List<PropertySurface> Properties, List<string> Attributes,
    List<EnumMember> EnumMembers);

internal record PackageSurface(string Id, List<TypeSurface> Types, List<TypeSurface> Activities);
