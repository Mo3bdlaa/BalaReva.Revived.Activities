using System.Activities;
using System.Data;
using BalaReva.Word.ContentEdit;
using BalaReva.Word.Documents;
using BalaReva.Word.ImageDocument;
using BalaReva.Word.PageHeaderFooter;
using BalaReva.Word.Pages;
using BalaReva.Word.Readers;
using BalaReva.Word.Tables;
using BalaReva.Word.Tools;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Tests;

/// <summary>
/// Each in-scope activity runs inside a real WordScope over a stand-in session, so the
/// scope-to-child handle contract is covered along with argument and output mapping.
/// </summary>
public class ActivityTests
{
    [Fact]
    public void The_scope_opens_the_document_and_closes_it_afterwards()
    {
        var service = new FakeWordService();

        Harness.Run(new AddPageBreak(), service, @"C:\docs\report.docx");

        Assert.Equal([@"Open(C:\docs\report.docx)", "AddPageBreak"], service.Calls);
        Assert.True(service.Disposed);
    }

    [Fact]
    public void The_scope_passes_its_passwords_to_the_session()
    {
        var service = new FakeWordService();
        var scope = new WordScope
        {
            FilePath = new InArgument<string>(@"C:\docs\secret.docx"),
            OpenPassword = new InArgument<string>("open-me"),
            ModifyPassword = new InArgument<string>("edit-me"),
        };

        var invoker = new WorkflowInvoker(scope);
        invoker.Extensions.Add(service);
        invoker.Invoke();

        Assert.Equal("open-me", service.Target.Password);
        Assert.Equal("edit-me", service.Target.ModiPassword);
    }

    [Fact]
    public void A_child_outside_a_scope_says_so()
    {
        var activity = new TableCount
        {
            ContinueOnError = new InArgument<bool>(false),
            Delay = new InArgument<double>(0),
            ExecutionResult = new OutArgument<bool>(),
            Result = new OutArgument<int>(),
        };

        var error = Assert.ThrowsAny<Exception>(() => WorkflowInvoker.Invoke(activity));
        Assert.Contains("Word Scope", Unwrap(error).Message, StringComparison.Ordinal);
    }

    [Fact]
    public void The_scope_still_closes_the_document_when_a_child_faults()
    {
        // Otherwise a failing child would leave a Word process running.
        var service = new FakeWordService { Throw = new InvalidOperationException("boom") };

        Assert.ThrowsAny<Exception>(() => Harness.Run(new AddPageBreak(), service));

        Assert.True(service.Disposed);
    }

    [Fact]
    public void ContinueOnError_reports_failure_instead_of_faulting()
    {
        var service = new FakeWordService { Throw = new InvalidOperationException("boom") };
        var run = new ScopeRun();
        var activity = new AddPageBreak
        {
            ContinueOnError = new InArgument<bool>(true),
            ExecutionResult = run.Capture<bool>("ExecutionResult"),
        };

        var outputs = run.Invoke(activity, service);

        Assert.False((bool)outputs.Values["ExecutionResult"]);
    }

    [Fact]
    public void TableCount_and_TableInfo_report_what_the_session_returns()
    {
        var service = new FakeWordService
        {
            Tables = 5,
            Info = new WordTableInfo { TotalRows = 4, TotalColumns = 3, HasHeaderRow = true },
        };

        var countRun = new ScopeRun();
        var count = countRun.Invoke(new TableCount { Result = countRun.Capture<int>("Result") }, service);

        var infoRun = new ScopeRun();
        var info = infoRun.Invoke(
            new TableInfo
            {
                TableIndex = new InArgument<int>(2),
                TotalRows = infoRun.Capture<int>("TotalRows"),
                TotalColumns = infoRun.Capture<int>("TotalColumns"),
                HasHeaderRow = infoRun.Capture<bool>("HasHeaderRow"),
            },
            service);

        Assert.Equal(5, count.Values["Result"]);
        Assert.Equal(4, info.Values["TotalRows"]);
        Assert.Equal(3, info.Values["TotalColumns"]);
        Assert.True((bool)info.Values["HasHeaderRow"]);
    }

    [Fact]
    public void WordStatistics_maps_every_count_onto_its_own_argument()
    {
        var service = new FakeWordService
        {
            Stats = new WordStatisticsResult
            {
                Pages = 12, WordCount = 3400, Lines = 500,
                Paragraphs = 90, Characters = 15000, CharactersWithSpaces = 18000,
            },
        };

        var run = new ScopeRun();
        var outputs = run.Invoke(
            new WordStatistics
            {
                Pages = run.Capture<long>("Pages"),
                WordCount = run.Capture<long>("WordCount"),
                Lines = run.Capture<long>("Lines"),
                Paragraphs = run.Capture<long>("Paragraphs"),
                Characters = run.Capture<long>("Characters"),
                CharactersWithSpaces = run.Capture<long>("CharactersWithSpaces"),
            },
            service);

        Assert.Equal(12L, outputs.Values["Pages"]);
        Assert.Equal(3400L, outputs.Values["WordCount"]);
        Assert.Equal(500L, outputs.Values["Lines"]);
        Assert.Equal(90L, outputs.Values["Paragraphs"]);
        Assert.Equal(15000L, outputs.Values["Characters"]);
        Assert.Equal(18000L, outputs.Values["CharactersWithSpaces"]);
    }

    [Fact]
    public void FindReplace_forwards_every_option()
    {
        var service = new FakeWordService();

        Harness.Run(
            new FindReplace
            {
                FindText = new InArgument<string>("old"),
                ReplaceText = new InArgument<string>("new"),
                FindOption = EnumFindReplace.Whole,
                ReplaceOption = EnumReplaceOption.Once,
                MatchCase = true,
            },
            service);

        Assert.Contains("FindReplace(old->new,Whole,Once,case=True)", service.Calls);
    }

    [Fact]
    public void SetTableValue_carries_the_font_flags_across()
    {
        var service = new FakeWordService();

        Harness.Run(
            new SetTableValue
            {
                TableIndex = new InArgument<int>(1),
                RowIndex = new InArgument<int>(2),
                ColumnIndex = new InArgument<int>(3),
                TextValue = new InArgument<string>("total"),
                FontBold = EnumSelectBoolean.True,
                FontItalic = EnumSelectBoolean.False,
                TextVerticalAlignment = EnumCellVerticalAlignment.VerticalCenter,
            },
            service);

        var format = Assert.IsType<WordCellFormat>(service.LastCellFormat);
        Assert.Equal("total", format.TextValue);
        Assert.Equal(EnumSelectBoolean.True, format.Bold);
        Assert.Equal(EnumSelectBoolean.False, format.Italic);
        // Untouched flags stay Select, which the service reads as "leave alone".
        Assert.Equal(EnumSelectBoolean.Select, format.Underline);
        Assert.Equal(EnumCellVerticalAlignment.VerticalCenter, format.VerticalAlignment);
    }

    [Fact]
    public void TableStyle_carries_the_banding_options_across()
    {
        var service = new FakeWordService();

        Harness.Run(
            new TableStyle
            {
                TableIndex = new InArgument<int>(1),
                StyleName = new InArgument<string>("Grid Table 4"),
                HeaderRow = true,
                BandedRows = true,
            },
            service);

        var options = Assert.IsType<WordTableStyleOptions>(service.LastStyleOptions);
        Assert.True(options.HeaderRow);
        Assert.True(options.BandedRows);
        Assert.False(options.TotalRow);
        Assert.Contains("TableStyle(1,Grid Table 4)", service.Calls);
    }

    [Fact]
    public void RowHeight_passes_an_empty_index_list_through_as_every_row()
    {
        var service = new FakeWordService();

        Harness.Run(
            new RowHeight
            {
                TableIndex = new InArgument<int>(1),
                tblRowHeight = new InArgument<float>(20),
                HeightRule = EnumRowHeightRule.Exactly,
            },
            service);

        Assert.Contains("RowHeight(1,[],20,Exactly)", service.Calls);
    }

    [Fact]
    public void InsertDataTable_passes_the_data_and_the_header_flag()
    {
        var service = new FakeWordService();
        var data = new DataTable();
        data.Columns.Add("Name");
        data.Rows.Add("Ada");
        data.Rows.Add("Grace");

        Harness.Run(
            new InsertDataTable
            {
                InputTable = new InArgument<DataTable>(_ => data),
                AddHeader = true,
                StyleName = new InArgument<string>("Grid"),
            },
            service);

        Assert.Contains("InsertDataTable(Bookmark,rows=2,header=True,Grid)", service.Calls);
    }

    [Fact]
    public void The_reader_activities_report_both_shapes()
    {
        var service = new FakeWordService();

        var fontRun = new ScopeRun();
        var byFont = fontRun.Invoke(
            new ReadByFont
            {
                FontStyle = EnumBoldItalicUnderline.Bold,
                ResultArray = fontRun.Capture<string[]>("ResultArray"),
                ResultTable = fontRun.Capture<DataTable>("ResultTable"),
            },
            service);

        var styleRun = new ScopeRun();
        var byStyle = styleRun.Invoke(
            new ReadByStyle
            {
                ParagraphStyle = new InArgument<string>("Heading 1"),
                ResultArray = styleRun.Capture<string[]>("ResultArray"),
            },
            service);

        Assert.Equal(["bold text"], (string[])byFont.Values["ResultArray"]);
        Assert.Single(((DataTable)byFont.Values["ResultTable"]).Rows);
        Assert.Equal(["heading"], (string[])byStyle.Values["ResultArray"]);
        Assert.Contains("ReadByStyle(Heading 1)", service.Calls);
    }

    [Fact]
    public void ReadHeaderFooter_reports_what_the_session_returns()
    {
        var service = new FakeWordService { HeaderFooterText = ["one", "two"] };

        var run = new ScopeRun();
        var outputs = run.Invoke(
            new ReadHeaderFooter
            {
                ReadType = EnumHeadersFooters.Footer,
                Result = run.Capture<string[]>("Result"),
            },
            service);

        Assert.Equal(["one", "two"], (string[])outputs.Values["Result"]);
        Assert.Contains("ReadHeaderFooter(Footer)", service.Calls);
    }

    [Fact]
    public void InsertHeaderFooter_gathers_the_formatting_into_one_request()
    {
        var service = new FakeWordService();

        Harness.Run(
            new InsertHeaderFooter
            {
                InsertType = EnumHeadersFooters.Header,
                OddPageText = new InArgument<string>("Report"),
                TextFontName = new InArgument<string>("Calibri"),
                FontSize = new InArgument<float>(14),
                FontBold = EnumBoolean.True,
                TextAlignment = HeaderFooterParagraphAlignment.Center,
            },
            service);

        var text = Assert.IsType<HeaderFooterText>(service.LastHeaderFooter);
        Assert.Equal("Report", text.OddPageText);
        Assert.Equal("Calibri", text.FontName);
        Assert.Equal(14, text.FontSize);
        Assert.Equal(EnumBoolean.True, text.Bold);
        Assert.Equal(HeaderFooterParagraphAlignment.Center, text.Alignment);
    }

    [Fact]
    public void ExecuteMacro_forwards_the_name_and_returns_the_result()
    {
        var service = new FakeWordService { MacroResult = "done" };

        var run = new ScopeRun();
        var outputs = run.Invoke(
            new ExecuteMacro
            {
                MacroName = new InArgument<string>("Module1.Run"),
                Parameters = new InArgument<object[]>(_ => new object[] { 1, "x" }),
                MacroOutput = run.Capture<object>("MacroOutput"),
            },
            service);

        Assert.Equal("done", outputs.Values["MacroOutput"]);
        Assert.Contains("ExecuteMacro(Module1.Run,[1,x])", service.Calls);
    }

    [Fact]
    public void FindAndReplace_really_does_change_passwords()
    {
        // Named wrong in the published package: its only arguments are the two new
        // passwords. Renaming it would break workflows, so it keeps the name.
        var service = new FakeWordService();

        Harness.Run(
            new FindAndReplace
            {
                NewOpenPassword = new InArgument<string>("open"),
                NewModifyPassword = new InArgument<string>("edit"),
            },
            service);

        Assert.Contains("ChangePassword(open,edit)", service.Calls);
    }

    [Fact]
    public void SaveAs_and_Print_and_ImageExtract_forward_their_options()
    {
        var service = new FakeWordService();

        Harness.Run(
            new SaveAs
            {
                NewFileName = new InArgument<string>(@"C:\out\copy.pdf"),
                SaveAsFormat = EnumSaveAs.PDF,
            },
            service);
        Harness.Run(
            new PrintDocument
            {
                PrinterName = new InArgument<string>("HP"),
                NoOfCopies = new InArgument<int>(3),
                Orientation = EnumOrientation.Landscape,
            },
            service);
        Harness.Run(
            new ImageExtract
            {
                ImageFolder = new InArgument<string>(@"C:\out\images"),
                FileExtension = EnumFileExtension.Jpeg,
            },
            service);

        Assert.Contains(@"SaveAs(C:\out\copy.pdf,PDF)", service.Calls);
        Assert.Contains("Print(HP,3,Landscape)", service.Calls);
        Assert.Contains(@"ExtractImages(C:\out\images,Jpeg)", service.Calls);
    }

    [Fact]
    public void ContentSelectionByBookMark_forwards_the_destination_and_the_extent()
    {
        var service = new FakeWordService();

        Harness.Run(
            new ContentSelectionByBookMark
            {
                GoTo = EnumGoTo.Page,
                GoToPageLineNo = new InArgument<int>(4),
                NoRows = new InArgument<int>(2),
                NoColumns = new InArgument<int>(3),
            },
            service);

        Assert.Contains("Select(Page,,4,2x3)", service.Calls);
    }

    [Fact]
    public void CloseAllWord_reaches_the_session()
    {
        var service = new FakeWordService();

        Harness.Run(new CloseAllWord(), service);

        Assert.Contains("CloseAllWord", service.Calls);
    }

    [Fact]
    public void MergeDocuments_runs_without_a_scope()
    {
        var service = new FakeWordService();

        Harness.RunStandalone(
            new Documents.MergeDocuments
            {
                WordFile = new InArgument<string>(@"C:\docs\main.docx"),
                SourceFiles = new InArgument<string[]>(_ => new[] { @"C:\docs\a.docx", @"C:\docs\b.docx" }),
            },
            service);

        Assert.Equal([@"Merge(C:\docs\main.docx,[C:\docs\a.docx,C:\docs\b.docx])"], service.Calls);
    }

    [Fact]
    public void MergeDocuments_rejects_an_empty_source_list()
    {
        var service = new FakeWordService();
        var activity = new Documents.MergeDocuments
        {
            WordFile = new InArgument<string>(@"C:\docs\main.docx"),
            SourceFiles = new InArgument<string[]>(_ => System.Array.Empty<string>()),
        };

        Assert.ThrowsAny<Exception>(() => Harness.RunStandalone(activity, service));
        Assert.Empty(service.Calls);
    }

    [Fact]
    public void WordToPdf_runs_without_a_scope_and_forwards_its_page_range()
    {
        var service = new FakeWordService();

        Harness.RunStandalone(
            new WordToPdf
            {
                WordFile = new InArgument<string>(@"C:\docs\main.docx"),
                PDFFile = new InArgument<string>(@"C:\out\main.pdf"),
                StartPage = new InArgument<int>(2),
                EndPage = new InArgument<int>(5),
            },
            service);

        Assert.Equal([@"WordToPdf(C:\docs\main.docx,C:\out\main.pdf,2-5)"], service.Calls);
    }

    private static Exception Unwrap(Exception error) =>
        error.InnerException is null ? error : Unwrap(error.InnerException);
}
