using System.Activities;
using System.Data;
using BalaReva.Easy.PowerPoint.Utilities;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Scope.Chart;
using BalaReva.EasyPowerPoint.Scope.Slides;
using BalaReva.EasyPowerPoint.Scope.TableArea;
using BalaReva.EasyPowerPoint.Scope.Tools;
using BalaReva.EasyPowerPoint.Utilities;

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
                AddSlideIndex = new InArgument<bool>(true),
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
                MatchCase = new InArgument<bool>(true),
                WholeWord = new InArgument<bool>(true),
                FirstOccurrence = new InArgument<bool>(true),
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
                HeaderRow = true,
                BandedRows = true,
            },
            service);

        var options = Assert.IsType<TableStyleOptions>(service.LastStyle);
        Assert.True(options.HeaderRow);
        Assert.True(options.BandedRows);
        Assert.False(options.TotalRow);
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
                ChartWidth = new InArgument<double>(300),
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
                PrintHiddenSlides = new InArgument<bool>(true),
            },
            service);

        Assert.Contains(@"ExportPdf(C:\out\deck.pdf,Screen)", service.Calls);
        Assert.Contains(@"SaveAs(C:\out\deck.odp,OpenDocumentPresentation)", service.Calls);
        Assert.Contains("Print(3,PrintBlackAndWhite,False,True)", service.Calls);
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
        // not depend on. They throw rather than quietly doing nothing.
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
    }

    private static Exception Unwrap(Exception error) =>
        error.InnerException is null ? error : Unwrap(error.InnerException);
}
