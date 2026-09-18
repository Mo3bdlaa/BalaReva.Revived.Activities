using System.Text;
using ICSharpCode.SharpZipLib.Checksum;

namespace BalaReva.Zip.Tests;

/// <summary>
/// Writes a zip a byte at a time, so an entry name lands in the archive exactly as given.
/// </summary>
/// <remarks>
/// A normal writer will not produce these archives. SharpZipLib's ZipOutputStream strips a
/// leading slash, rewrites backslashes and drops a UNC prefix as it writes, so an archive
/// built with it cannot carry <c>/etc/passwd</c> — it comes back as <c>etc/passwd</c>,
/// which is harmless. An attacker is not using a polite writer, so neither is this: the
/// local header and central directory are assembled here, and whatever string is asked
/// for is what an extractor will read back.
/// </remarks>
internal static class HostileZip
{
    /// <summary>Writes a stored-method zip holding the given entries, names untouched.</summary>
    public static string Write(string path, params (string Entry, string Content)[] entries)
    {
        using var stream = File.Create(path);
        var central = new List<byte[]>();
        var offsets = new List<uint>();

        foreach (var (entry, content) in entries)
        {
            var name = Encoding.UTF8.GetBytes(entry);
            var data = Encoding.UTF8.GetBytes(content);

            var crc = new Crc32();
            crc.Update(data);
            var checksum = (uint)crc.Value;

            offsets.Add((uint)stream.Position);

            // Local file header: stored, no data descriptor, sizes known up front.
            Write(stream, 0x04034B50u);            // signature
            Write(stream, (ushort)20);             // version needed
            Write(stream, (ushort)0x0800);         // flags: UTF-8 names
            Write(stream, (ushort)0);              // method: stored
            Write(stream, (ushort)0);              // modification time
            Write(stream, (ushort)0x21);           // modification date: 1980-01-01
            Write(stream, checksum);
            Write(stream, (uint)data.Length);      // compressed size
            Write(stream, (uint)data.Length);      // uncompressed size
            Write(stream, (ushort)name.Length);
            Write(stream, (ushort)0);              // extra field length
            stream.Write(name);
            stream.Write(data);

            central.Add(CentralHeader(name, checksum, data.Length, offsets[^1]));
        }

        var directoryStart = (uint)stream.Position;
        foreach (var header in central) stream.Write(header);
        var directorySize = (uint)stream.Position - directoryStart;

        // End of central directory.
        Write(stream, 0x06054B50u);
        Write(stream, (ushort)0);                  // this disk
        Write(stream, (ushort)0);                  // disk with the directory
        Write(stream, (ushort)entries.Length);     // entries on this disk
        Write(stream, (ushort)entries.Length);     // entries in total
        Write(stream, directorySize);
        Write(stream, directoryStart);
        Write(stream, (ushort)0);                  // comment length

        return path;
    }

    private static byte[] CentralHeader(byte[] name, uint crc, int size, uint offset)
    {
        using var buffer = new MemoryStream();
        Write(buffer, 0x02014B50u);                // signature
        Write(buffer, (ushort)20);                 // version made by
        Write(buffer, (ushort)20);                 // version needed
        Write(buffer, (ushort)0x0800);             // flags: UTF-8 names
        Write(buffer, (ushort)0);                  // method: stored
        Write(buffer, (ushort)0);                  // modification time
        Write(buffer, (ushort)0x21);               // modification date
        Write(buffer, crc);
        Write(buffer, (uint)size);                 // compressed size
        Write(buffer, (uint)size);                 // uncompressed size
        Write(buffer, (ushort)name.Length);
        Write(buffer, (ushort)0);                  // extra field length
        Write(buffer, (ushort)0);                  // comment length
        Write(buffer, (ushort)0);                  // disk number start
        Write(buffer, (ushort)0);                  // internal attributes
        Write(buffer, 0u);                         // external attributes
        Write(buffer, offset);
        buffer.Write(name);
        return buffer.ToArray();
    }

    private static void Write(Stream stream, uint value) =>
        stream.Write(BitConverter.GetBytes(value));

    private static void Write(Stream stream, ushort value) =>
        stream.Write(BitConverter.GetBytes(value));
}
