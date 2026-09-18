using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BalaReva.EasyExcel.Tests;

/// <summary>
/// The one part of this package that can be tested against real files: the Open XML
/// reader behind GetHiddenRows and GetHiddenColumns needs no Excel installation.
/// </summary>
public sealed class OpenXmlReaderTests : IDisposable
{
    private readonly string _folder =
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName())).FullName;

    private readonly IOpenXmlReader _reader = new OpenXmlReader();

    [Fact]
    public void Hidden_rows_come_back_by_number()
    {
        var path = Write("book.xlsx", hiddenRows: [2, 5], hiddenColumns: []);

        Assert.Equal([2, 5], _reader.HiddenRows(path, "Data"));
    }

    [Fact]
    public void Hidden_columns_come_back_by_letter()
    {
        // B:D as one span and G on its own, which is how Excel writes a hidden run.
        var path = Write("book.xlsx", hiddenRows: [], hiddenColumns: [(2, 4), (7, 7)]);

        Assert.Equal(["B", "C", "D", "G"], _reader.HiddenColumns(path, "Data"));
    }

    [Fact]
    public void A_sheet_with_nothing_hidden_comes_back_empty()
    {
        var path = Write("book.xlsx", hiddenRows: [], hiddenColumns: []);

        Assert.Empty(_reader.HiddenRows(path, "Data"));
        Assert.Empty(_reader.HiddenColumns(path, "Data"));
    }

    [Fact]
    public void An_empty_sheet_name_reads_the_first_sheet()
    {
        var path = Write("book.xlsx", hiddenRows: [3], hiddenColumns: []);

        Assert.Equal([3], _reader.HiddenRows(path, string.Empty));
    }

    [Fact]
    public void A_sheet_that_is_not_there_says_so()
    {
        var path = Write("book.xlsx", hiddenRows: [], hiddenColumns: []);

        var error = Assert.Throws<ArgumentException>(() => _reader.HiddenRows(path, "Nowhere"));
        Assert.Contains("Nowhere", error.Message);
    }

    [Fact]
    public void A_file_that_is_not_there_says_so() =>
        Assert.Throws<FileNotFoundException>(
            () => _reader.HiddenRows(Path.Combine(_folder, "missing.xlsx"), "Data"));

    [Fact]
    public void The_reader_does_not_hold_the_file_open()
    {
        var path = Write("book.xlsx", hiddenRows: [2], hiddenColumns: []);
        _reader.HiddenRows(path, "Data");

        // Reading detaches the worksheet from the package before disposing it; if it did
        // not, this delete would fail on Windows and the activity would lock the file.
        File.Delete(path);
        Assert.False(File.Exists(path));
    }

    /// <summary>Writes a one-sheet workbook with the given rows and column spans hidden.</summary>
    private string Write(string name, uint[] hiddenRows, (uint First, uint Last)[] hiddenColumns)
    {
        var path = Path.Combine(_folder, name);
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);

        var book = document.AddWorkbookPart();
        book.Workbook = new Workbook();

        var sheetPart = book.AddNewPart<WorksheetPart>();
        var data = new SheetData();

        for (uint row = 1; row <= 8; row++)
        {
            var element = new Row { RowIndex = row };
            if (hiddenRows.Contains(row)) element.Hidden = true;
            data.Append(element);
        }

        var worksheet = new Worksheet();
        if (hiddenColumns.Length > 0)
        {
            var columns = new Columns();
            foreach (var (first, last) in hiddenColumns)
                columns.Append(new Column { Min = first, Max = last, Hidden = true, CustomWidth = true });
            worksheet.Append(columns);
        }
        worksheet.Append(data);
        sheetPart.Worksheet = worksheet;

        // Fully qualified: BalaReva.EasyExcel.Sheets is a namespace of this package, so
        // the bare name resolves to that instead.
        book.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets()).Append(new Sheet
        {
            Id = book.GetIdOfPart(sheetPart),
            SheetId = 1,
            Name = "Data",
        });

        book.Workbook.Save();
        return path;
    }

    public void Dispose()
    {
        try { Directory.Delete(_folder, recursive: true); }
        catch (IOException) { /* the agent will clean its own temp */ }
    }
}
