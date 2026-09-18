using System.Activities;
using System.Data;
using BalaReva.Easy.PowerPoint.Utilities;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Scope.Chart;
using BalaReva.EasyPowerPoint.Scope.Slides;
using BalaReva.EasyPowerPoint.Scope.TableArea;
using BalaReva.EasyPowerPoint.Scope.Tools;
using BalaReva.EasyPowerPoint.Scope.Main;
using BalaReva.EasyPowerPoint.Utilities;
using BalaReva.PowerPoint;

namespace BalaReva.EasyPowerPoint.Tests;

/// <summary>
/// Each activity runs inside a real PowerPointScope over a stand-in presentation, so
/// the scope-to-child handle contract is covered along with argument and output mapping.
/// </summary>
public class ActivityTests
{
    [Fact]
    public void The_scope_opens_the_presentation_and_closes_it_afterwards()
    {
        var service = new FakePowerPointService();

        Harness.Run(new UpdateLinks(), service, @"C:\decks\q1.pptx");

        Assert.Equal([@"Open(C:\decks\q1.pptx)", "UpdateLinks"], service.Calls);
        Assert.True(service.Disposed);
    }

    [Fact]
    public void The_scope_passes_its_macro_setting_through()
    {
        var service = new FakePowerPointService();

        new ScopeRun().Invoke(new UpdateLinks(), service, macros: EnableDisableEnum.Enable);

        Assert.True(service.MacrosWereEnabled);
    }

    [Fact]
    public void A_child_outside_a_scope_says_so()
    {
        var activity = new UpdateLinks
        {
            ContinueOnError = new InArgument<bool>(false),
            Delay = new InArgument<double>(0),
            ExecutionResult = new OutArgument<bool>(),
        };

        var error = Assert.ThrowsAny<Exception>(() => WorkflowInvoker.Invoke(activity));
        Assert.Contains("PowerPoint Scope", Unwrap(error).Message, StringComparison.Ordinal);
    }

    [Fact]
    public void The_scope_still_closes_the_presentation_when_a_child_faults()
    {
        var service = new FakePowerPointService { Throw = new InvalidOperationException("boom") };

        Assert.ThrowsAny<Exception>(() => Harness.Run(new UpdateLinks(), service));

        Assert.True(service.Disposed);
    }

    [Fact]
    public void ContinueOnError_reports_failure_instead_of_faulting()
    {
        var service = new FakePowerPointService { Throw = new InvalidOperationException("boom") };
        var run = new ScopeRun();
        var activity = new UpdateLinks
        {
            ContinueOnError = new InArgument<bool>(true),
            ExecutionResult = run.Capture<bool>("ExecutionResult"),
        };

        var outputs = run.Invoke(activity, service);

        Assert.False((bool)outputs.Values["ExecutionResult"]);
    }

    [Fact]
    public void SlideCount_reports_what_the_presentation_returns()
    {
        var service = new FakePowerPointService { Slides = 17 };
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new SlideCount { TotalSlides = run.Capture<int>("TotalSlides") }, service);

        Assert.Equal(17, outputs.Values["TotalSlides"]);
    }

    [Fact]
    public void The_count_activities_keep_the_published_swapped_output_names()
    {
        // ImageShapeCount reports through ChartCount, TextShapeCount through ImageCount.
        var service = new FakePowerPointService();

        var imageRun = new ScopeRun();
        var images = imageRun.Invoke(
            new ImageShapeCount
            {
                SlideIndex = new InArgument<int>(1),
                ChartCount = imageRun.Capture<int>("ChartCount"),
            },
            service);

        var textRun = new ScopeRun();
        var text = textRun.Invoke(
            new TextShapeCount
            {
                SlideIndex = new InArgument<int>(1),
                ImageCount = textRun.Capture<int>("ImageCount"),
            },
            service);

        Assert.Equal(5, images.Values["ChartCount"]);
        Assert.Equal(7, text.Values["ImageCount"]);
    }

    [Fact]
    public void ReadText_maps_both_shapes_of_its_result()
    {
        var service = new FakePowerPointService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new ReadText
            {
                SlideIndexes = new InArgument<int[]>(_ => new[] { 1, 2 }),
                AddSlideIndex = true,
                ResultArray = run.Capture<string[]>("ResultArray"),
                ResultString = run.Capture<string>("ResultString"),
            },
            service);

        Assert.Equal(["one", "two"], (string[])outputs.Values["ResultArray"]);
        Assert.Equal("one\ntwo", outputs.Values["ResultString"]);
        Assert.Contains("ReadText([1,2],True,False)", service.Calls);
    }

    [Fact]
    public void An_empty_slide_list_is_passed_through_as_every_slide()
    {
        var service = new FakePowerPointService();
        var run = new ScopeRun();

        run.Invoke(
            new ReadText
            {
                ResultArray = run.Capture<string[]>("ResultArray"),
                ResultString = run.Capture<string>("ResultString"),
            },
            service);

        // Choosing what an empty list means is the presentation's job, not the activity's.
        Assert.Contains("ReadText([],False,False)", service.Calls);
    }

    [Fact]
    public void Find_Replace_forwards_every_option()
    {
        var service = new FakePowerPointService();

        Harness.Run(
            new Find_Replace
            {
                FindText = new InArgument<string>("old"),
                ReplaceText = new InArgument<string>("new"),
                MatchCase = true,
                WholeWord = true,
                FirstOccurrence = true,
            },
            service);

        Assert.Contains("FindReplace([],old->new,True,True,first=True)", service.Calls);
    }

    [Fact]
    public void InsertTextBox_gathers_its_font_arguments_into_the_request()
    {
        var service = new FakePowerPointService();

        Harness.Run(
            new InsertTextBox
            {
                SlideIndex = new InArgument<int>(2),
                Text = new InArgument<string>("Title"),
                FontName = new InArgument<string>("Calibri"),
                FontSize = new InArgument<float>(28),
                FontBold = TrueFalseNoneEnum.True,
                TextAlignment = EnumTextEffectAlignment.Centered,
            },
            service);

        var request = Assert.IsType<TextBoxRequest>(service.LastTextBox);
        Assert.Equal("Title", request.Text);
        Assert.Equal("Calibri", request.Style.FontName);
        Assert.Equal(28, request.Style.FontSize);
        Assert.Equal(TrueFalseNoneEnum.True, request.Style.Bold);
        // Untouched flags stay None, which the service reads as "leave it alone".
        Assert.Equal(TrueFalseNoneEnum.None, request.Style.Italic);
        Assert.Equal(EnumTextEffectAlignment.Centered, request.Alignment);
    }

    [Fact]
    public void InsertPicture_passes_zero_sizes_through_as_keep_the_original()
    {
        var service = new FakePowerPointService();

        Harness.Run(
            new InsertPicture
            {
                SlideIndex = new InArgument<int>(1),
                ImagePath = new InArgument<string>(@"C:\img\logo.png"),
                ImageLeft = new InArgument<float>(10),
            },
            service);

        var request = Assert.IsType<PictureRequest>(service.LastPicture);
        Assert.Equal(@"C:\img\logo.png", request.ImagePath);
        Assert.Equal(10, request.Left);
        Assert.Equal(0, request.Width);
        Assert.Equal(0, request.Height);
    }

    [Fact]
    public void A_table_activity_prefers_the_name_over_the_index()
    {
        var service = new FakePowerPointService();

        Harness.Run(
            new DeleteTable
            {
                SlideIndex = new InArgument<int>(3),
                TableIndex = new InArgument<int>(1),
                TableName = new InArgument<string>("Sales"),
            },
            service);

        var reference = Assert.IsType<TableRef>(service.LastTableRef);
        Assert.Equal("Sales", reference.TableName);
        Assert.Equal(3, reference.SlideIndex);
        // Both are carried; choosing between them is the presentation's job.
        Assert.Equal(1, reference.TableIndex);
    }

    [Fact]
    public void GetRowItem_reports_the_cell_text()
    {
        var service = new FakePowerPointService { CellText = "hello" };
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new BalaReva.Easy.PowerPoint.Scope.TableArea.GetRowItem
            {
                SlideIndex = new InArgument<int>(1),
                TableIndex = new InArgument<int>(1),
                RowIndex = new InArgument<int>(2),
                ColumnIndex = new InArgument<int>(3),
                Value = run.Capture<string>("Value"),
            },
            service);

        Assert.Equal("hello", outputs.Values["Value"]);
    }

    [Fact]
    public void AddTable_rejects_a_missing_input_table()
    {
        var service = new FakePowerPointService();
        var activity = new AddTable { SlideIndex = new InArgument<int>(1) };

        Assert.ThrowsAny<Exception>(() => Harness.Run(activity, service));
    }

    [Fact]
    public void AddTable_carries_its_data_and_emphasis_options()
    {
        var service = new FakePowerPointService();
        var data = new DataTable();
        data.Columns.Add("Region");
        data.Rows.Add("North");
        data.Rows.Add("South");

        Harness.Run(
            new AddTable
            {
                SlideIndex = new InArgument<int>(1),
                InputTable = new InArgument<DataTable>(_ => data),
                TableName = new InArgument<string>("Regions"),
                AddHeader = true,
                FirstColumn = true,
            },
            service);

        var request = Assert.IsType<AddTableRequest>(service.LastTable);
        Assert.Equal("Regions", request.TableName);
        Assert.True(request.AddHeader);
        Assert.True(request.FirstColumn);
        Assert.False(request.LastColumn);
        Assert.Equal(2, request.InputTable.Rows.Count);
    }

    [Fact]
    public void StyleOption_carries_the_banding_options()
    {
        var service = new FakePowerPointService();

        Harness.Run(
            new StyleOption
            {
                SlideIndex = new InArgument<int>(1),
                TableName = new InArgument<string>("Sales"),
                HeaderRow = TrueFalseNoneEnum.True,
                BandedRows = TrueFalseNoneEnum.True,
            },
            service);

        var options = Assert.IsType<TableStyleOptions>(service.LastStyle);
        Assert.Equal(TrueFalseNoneEnum.True, options.HeaderRow);
        Assert.Equal(TrueFalseNoneEnum.True, options.BandedRows);
        // Left unset, so it stays None: the table keeps whatever it already had.
        Assert.Equal(TrueFalseNoneEnum.None, options.TotalRow);
    }

    [Fact]
    public void The_chart_activities_forward_their_indexes()
    {
        var service = new FakePowerPointService();

        Harness.Run(
            new ChartDelete
            {
                SlideIndex = new InArgument<int>(2),
                ChartIndex = new InArgument<int>(1),
            },
            service);
        Harness.Run(
            new ChartFormat
            {
                SlideIndex = new InArgument<int>(2),
                ChartIndex = new InArgument<int>(1),
                ChartWidth = new InArgument<float>(300),
            },
            service);

        Assert.Contains("ChartDelete(2,1)", service.Calls);
        Assert.Contains("ChartFormat(2,1,0,0,300,0)", service.Calls);
    }

    [Fact]
    public void ExportPdf_and_SaveAs_and_Print_forward_their_options()
    {
        var service = new FakePowerPointService();

        Harness.Run(
            new ExportPdf
            {
                FilePath = new InArgument<string>(@"C:\out\deck.pdf"),
                FormatType = FixedFormatIntentEnum.Screen,
            },
            service);
        Harness.Run(
            new BalaReva.EasyPowerPoint.Scope.Slides.SaveAs
            {
                SaveAsFile = new InArgument<string>(@"C:\out\deck.odp"),
                SaveAsFormat = SaveAsEnum.OpenDocumentPresentation,
            },
            service);
        Harness.Run(
            new BalaReva.EasyPowerPoint.Scope.Slides.Print
            {
                NumberOfCopies = new InArgument<int>(3),
                PrintColorType = PrintColorTypeEnum.PrintBlackAndWhite,
                PrintHiddenSlides = TrueFalseNoneEnum.True,
            },
            service);

        Assert.Contains(@"ExportPdf(C:\out\deck.pdf,Screen)", service.Calls);
        Assert.Contains(@"SaveAs(C:\out\deck.odp,OpenDocumentPresentation)", service.Calls);
        Assert.Contains("Print(3,PrintBlackAndWhite,None,True)", service.Calls);
    }

    [Fact]
    public void DataTransformer_rejects_a_missing_dictionary()
    {
        var service = new FakePowerPointService();
        var activity = new DataTransformer();

        Assert.ThrowsAny<Exception>(() => Harness.Run(activity, service));
    }

    [Fact]
    public void The_Excel_bridging_activities_say_plainly_that_they_are_not_implemented()
    {
        // Both would need the Excel object model, which this package deliberately does
        // not depend on. They throw rather than quietly doing nothing, and they do it in
        // the activity, so the refusal holds whichever service is behind them.
        var service = new FakePowerPointService();

        var export = Assert.ThrowsAny<Exception>(() => Harness.Run(
            new ExportTableToExcel
            {
                SlideIndex = new InArgument<int>(1),
                TableIndex = new InArgument<int>(1),
                ExcelFile = new InArgument<string>(@"C:\out\book.xlsx"),
            },
            service));

        Assert.Contains("not implemented", Unwrap(export).Message, StringComparison.OrdinalIgnoreCase);

        // The table still reaches the clipboard, which is the route it leaves open.
        Assert.Contains(service.Calls, call => call.StartsWith("TableCopyToClipboard"));

        var import = Assert.ThrowsAny<Exception>(() => Harness.Run(
            new ImportDataFromExcel
            {
                SlideIndex = new InArgument<int>(1),
                ExcelFile = new InArgument<string>(@"C:\out\book.xlsx"),
            },
            new FakePowerPointService()));

        Assert.Contains("not implemented", Unwrap(import).Message, StringComparison.OrdinalIgnoreCase);
    }

    // The activities below carry the property types the published package declared, which
    // are not the obvious ones: arrays where a single value would do, three-state enums
    // where a bool would do, and design-time properties where an argument would do. Each
    // of these got it wrong first time round, so each has a test.

    [Fact]
    public void The_scope_hands_its_body_the_file_it_opened()
    {
        var service = new FakePowerPointService();

        var scope = new PowerPointScope
        {
            FilePath = new InArgument<string>(@"C:\decks\deck.pptx"),
            OpenPassword = new InArgument<string>("open"),
            ModifyPassword = new InArgument<string>("modify"),
        };

        var argument = new DelegateInArgument<PowerPointObject> { Name = "PowerPointPresentation" };
        var captured = new Outputs();
        scope.Body = new ActivityAction<PowerPointObject>
        {
            Argument = argument,
            Handler = new CaptureOutput<PowerPointObject>(captured, "Body")
            {
                Value = new InArgument<PowerPointObject>(argument),
            },
        };

        var invoker = new WorkflowInvoker(scope);
        invoker.Extensions.Add(service);
        invoker.Invoke();

        var target = Assert.IsType<PowerPointObject>(captured.Values["Body"]);
        Assert.Equal(@"C:\decks\deck.pptx", target.FilePath);
        Assert.Equal("open", target.Password);
        Assert.Equal("modify", target.ModiPassword);
        // Null behind a stand-in: there is no COM presentation to hand over.
        Assert.Null(target.PptPersentation);
    }

    [Fact]
    public void SlideExtractor_reports_the_slide_text_shapes()
    {
        var service = new FakePowerPointService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new SlideExtractor
            {
                SlideIndex = new InArgument<int>(2),
                SlideResult = run.Capture<SlideObject>("SlideResult"),
            },
            service);

        var slide = Assert.IsType<SlideObject>(outputs.Values["SlideResult"]);
        var shape = Assert.Single(slide.TextShapes);
        Assert.Equal("slide 2", shape.Text);
        Assert.Equal(300, shape.BoundWidth);
        Assert.Equal("Calibri", shape.TextShapeFont.Name);
        Assert.True(shape.TextShapeFont.Bold);
    }

    [Fact]
    public void FindText_reports_the_slides_it_matched_on()
    {
        var service = new FakePowerPointService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new FindText
            {
                SlideIndexes = new InArgument<int[]>(_ => new[] { 1, 2, 5 }),
                FindString = new InArgument<string>("total"),
                MatchCase = true,
                WholeWord = true,
                ResultArray = run.Capture<int[]>("ResultArray"),
                ResultTable = run.Capture<DataTable>("ResultTable"),
            },
            service);

        Assert.Contains("FindText([1,2,5],total,True,True)", service.Calls);
        Assert.Equal(new[] { 2, 5 }, Assert.IsType<int[]>(outputs.Values["ResultArray"]));
    }

    [Fact]
    public void CommentsRead_reports_one_entry_per_comment()
    {
        var service = new FakePowerPointService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new CommentsRead
            {
                SlideIndex = new InArgument<int>(1),
                IncludeReplies = true,
                Result = run.Capture<string[]>("Result"),
                ResultTable = run.Capture<DataTable>("ResultTable"),
            },
            service);

        Assert.Contains("CommentsRead(1,replies=True)", service.Calls);
        Assert.Equal(
            new[] { "alice: nice", "bob: agreed" },
            Assert.IsType<string[]>(outputs.Values["Result"]));
    }

    [Fact]
    public void ReadTables_reads_several_slides_and_passes_its_header_flag_through()
    {
        var service = new FakePowerPointService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new ReadTables
            {
                SlideIndex = new InArgument<int[]>(_ => new[] { 1, 3 }),
                HasHeader = true,
                ResultSet = run.Capture<DataSet>("ResultSet"),
            },
            service);

        Assert.Contains("ExtractTables([1,3],header=True)", service.Calls);
        Assert.True(service.LastHasHeader);
        Assert.Equal(2, Assert.IsType<DataSet>(outputs.Values["ResultSet"]).Tables.Count);
    }

    [Fact]
    public void ExtractTables_reports_an_array_for_the_one_slide_it_is_given()
    {
        var service = new FakePowerPointService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new ExtractTables
            {
                SlideIndex = new InArgument<int>(4),
                OutputTables = run.Capture<DataTable[]>("OutputTables"),
            },
            service);

        Assert.Contains("ExtractTables([4],header=False)", service.Calls);
        Assert.Equal(2, Assert.IsType<DataTable[]>(outputs.Values["OutputTables"]).Length);
    }

    [Fact]
    public void RefreshData_takes_the_shorts_the_published_package_declared()
    {
        var service = new FakePowerPointService();

        Harness.Run(
            new RefreshData { SlideIndex = new InArgument<short[]>(_ => new short[] { 2, 4 }) },
            service);

        Assert.Contains("RefreshData([2,4])", service.Calls);
    }

    [Fact]
    public void TextShapeEdit_forwards_the_whole_shape_it_was_given()
    {
        var service = new FakePowerPointService();
        var wanted = new TextShape
        {
            Text = "Revenue",
            BoundLeft = 40,
            TextShapeFont = new ShapeFont { Name = "Arial", Size = 24, Underline = true },
        };

        Harness.Run(
            new TextShapeEdit
            {
                SlideIndex = new InArgument<int>(1),
                TextIndex = new InArgument<int>(2),
                TextStyle = new InArgument<TextShape>(_ => wanted),
            },
            service);

        Assert.Same(wanted, service.LastTextShape);
        Assert.Contains("TextShapeEdit(1,2)", service.Calls);
    }

    private static Exception Unwrap(Exception error) =>
        error.InnerException is null ? error : Unwrap(error.InnerException);
}
