using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BalaReva.EasyExcel;

/// <summary>
/// Reads hidden rows and columns straight out of the workbook's Open XML package.
/// </summary>
/// <remarks>
/// The only part of this package that does not need Excel. It is also the only part that
/// a Linux agent can execute, so it is the only part with tests that run against real
/// files rather than a stand-in.
/// </remarks>
public sealed class OpenXmlReader : IOpenXmlReader
{
    /// <summary>The reader the activities use when no extension is registered.</summary>
    public static IOpenXmlReader Instance { get; } = new OpenXmlReader();

    /// <inheritdoc />
    public List<string> HiddenColumns(string filePath, string sheetName)
    {
        var hidden = new List<string>();
        foreach (var column in Sheet(filePath, sheetName).Descendants<Column>())
        {
            if (column.Hidden?.Value != true) continue;

            var min = (int)(column.Min?.Value ?? 0);
            var max = (int)(column.Max?.Value ?? 0);

            // One Column element covers a span, so a single hidden run of B:D arrives
            // as one element rather than three.
            for (var index = min; index >= 1 && index <= max; index++)
                hidden.Add(ColumnName(index));
        }
        return hidden;
    }

    /// <inheritdoc />
    public List<int> HiddenRows(string filePath, string sheetName) =>
    [
        .. Sheet(filePath, sheetName).Descendants<Row>()
            .Where(row => row.Hidden?.Value == true && row.RowIndex is not null)
            .Select(row => (int)row.RowIndex!.Value)
            .Where(index => index > 0),
    ];

    /// <summary>The worksheet part for a sheet, or the first sheet when none is named.</summary>
    private static Worksheet Sheet(string filePath, string sheetName)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"'{filePath}' does not exist.", filePath);

        using var document = SpreadsheetDocument.Open(filePath, isEditable: false);
        var book = document.WorkbookPart
            ?? throw new InvalidOperationException($"'{filePath}' has no workbook part.");

        var sheets = book.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];
        var sheet = string.IsNullOrWhiteSpace(sheetName)
            ? sheets.FirstOrDefault()
            : sheets.FirstOrDefault(s =>
                string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));

        if (sheet?.Id?.Value is null)
            throw new ArgumentException(
                string.IsNullOrWhiteSpace(sheetName)
                    ? $"'{filePath}' has no sheets."
                    : $"'{filePath}' has no sheet called '{sheetName}'.",
                nameof(sheetName));

        var part = (WorksheetPart)book.GetPartById(sheet.Id.Value);
        var worksheet = part.Worksheet
            ?? throw new InvalidOperationException($"'{filePath}' has no worksheet part for that sheet.");

        // The document is disposed on the way out of this method, so the worksheet has to
        // be detached from it first. What comes back is row and column metadata, not data.
        return (Worksheet)worksheet.CloneNode(deep: true);
    }

    /// <summary>Turns a column number into its letters. Excel counts from 1, with no zero digit.</summary>
    private static string ColumnName(int index)
    {
        var name = string.Empty;
        while (index > 0)
        {
            var digit = (index - 1) % 26;
            name = (char)('A' + digit) + name;
            index = (index - 1) / 26;
        }
        return name;
    }
}
