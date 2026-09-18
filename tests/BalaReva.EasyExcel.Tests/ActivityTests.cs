using System.Activities;
using System.Data;
using System.Drawing;
using BalaReva.EasyExcel.AddIns;
using BalaReva.EasyExcel.Charts;
using BalaReva.EasyExcel.EmptyRows;
using BalaReva.EasyExcel.FormulaArea;
using BalaReva.EasyExcel.FreezePanes;
using BalaReva.EasyExcel.HyperLinks;
using BalaReva.EasyExcel.Main;
using BalaReva.EasyExcel.Outline;
using BalaReva.EasyExcel.Settings.TrustedLocation;
using BalaReva.EasyExcel.Sheet_Images;
using BalaReva.EasyExcel.Sheets;
using BalaReva.EasyExcel.Tables;
using BalaReva.EasyExcel.Tools;
using BalaReva.EasyExcel.Utilities;
using BalaReva.EasyExcel.WorkBook;

namespace BalaReva.EasyExcel.Tests;

/// <summary>
/// Each activity runs inside a real ExcelScope over a stand-in workbook, so the
/// scope-to-child handle contract is covered along with argument and output mapping.
/// </summary>
public class ActivityTests
{
    [Fact]
    public void The_scope_opens_the_workbook_and_closes_it_afterwards()
    {
        var service = new FakeExcelService();

        Harness.Run(new RefreshAll(), service, @"C:\books\ledger.xlsx");

        Assert.Contains(@"Open(C:\books\ledger.xlsx)", service.Calls);
        Assert.Contains("RefreshAll()", service.Calls);
        Assert.True(service.Disposed);
    }

    [Fact]
    public void The_scope_hands_its_body_the_file_it_opened()
    {
        var service = new FakeExcelService();
        var captured = new Outputs();

        var argument = new DelegateInArgument<ExcelParam> { Name = "ExcelWorkBook" };
        var scope = new ExcelScope
        {
            FilePath = new InArgument<string>(@"C:\books\ledger.xlsx"),
            FilePassword = new InArgument<string>("open"),
            ModifyPassword = new InArgument<string>("modify"),
            Body = new ActivityAction<ExcelParam>
            {
                Argument = argument,
                Handler = new CaptureOutput<ExcelParam>(captured, "Body")
                {
                    Value = new InArgument<ExcelParam>(argument),
                },
            },
        };

        var invoker = new WorkflowInvoker(scope);
        invoker.Extensions.Add(service);
        invoker.Invoke();

        var target = Assert.IsType<ExcelParam>(captured.Values["Body"]);
        Assert.Equal(@"C:\books\ledger.xlsx", target.FilePath);
        Assert.Equal("open", target.Password);
        Assert.Equal("modify", target.ModiPassword);
        // Null behind a stand-in: there is no COM workbook to hand over.
        Assert.Null(target.ExcelWorkBook);
    }

    [Fact]
    public void The_scope_still_closes_the_workbook_when_a_child_faults()
    {
        var service = new FakeExcelService { Throw = new InvalidOperationException("boom") };

        Assert.ThrowsAny<Exception>(() => Harness.Run(new RefreshAll(), service));
        Assert.True(service.Disposed);
    }

    [Fact]
    public void A_child_outside_a_scope_says_so()
    {
        var service = new FakeExcelService();
        var activity = new RefreshAll
        {
            ContinueOnError = new InArgument<bool>(false),
            Delay = new InArgument<short>(0),
            ExecutionResult = new OutArgument<bool>(),
        };

        var error = Assert.ThrowsAny<Exception>(() => Harness.RunAlone(activity, service));
        Assert.Contains("Excel Scope", Unwrap(error).Message);
    }

    [Fact]
    public void ContinueOnError_reports_failure_instead_of_faulting()
    {
        var service = new FakeExcelService { Throw = new InvalidOperationException("boom") };
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new RefreshAll
            {
                ContinueOnError = new InArgument<bool>(true),
                ExecutionResult = run.Capture<bool>("ExecutionResult"),
            },
            service);

        Assert.False((bool)outputs.Values["ExecutionResult"]);
    }

    [Fact]
    public void The_scope_passes_its_macro_setting_through()
    {
        var service = new FakeExcelService();
        new ScopeRun().Invoke(new RefreshAll(), service, macros: EnableDisableEnum.Enable);

        Assert.True(service.LastOpen!.MacrosEnabled);
    }

    [Fact]
    public void A_sheet_activity_forwards_the_sheet_it_was_given()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new SelectCell
            {
                SheetName = new InArgument<string>("Summary"),
                CellRange = new InArgument<string>("A1:D10"),
            },
            service);

        Assert.Equal("Summary", service.LastSheetName);
        Assert.Contains("SelectCell(Summary,A1:D10)", service.Calls);
    }

    [Fact]
    public void An_empty_sheet_name_is_passed_through_as_the_active_sheet()
    {
        var service = new FakeExcelService();

        Harness.Run(new RemoveFilterNonTable(), service);

        Assert.Equal(string.Empty, service.LastSheetName);
        Assert.Contains("RemoveFilterNonTable()", service.Calls);
    }

    [Fact]
    public void The_six_range_functions_each_ask_for_their_own_kind()
    {
        var service = new FakeExcelService();
        var range = new InArgument<string>("A1:A20");

        Harness.Run(new AverageRange { CellRange = range }, service);
        Harness.Run(new CountRange { CellRange = range }, service);
        Harness.Run(new CountARange { CellRange = range }, service);
        Harness.Run(new MaxRange { CellRange = range }, service);
        Harness.Run(new MinRange { CellRange = range }, service);
        Harness.Run(new SumRange { CellRange = range }, service);

        foreach (var kind in Enum.GetNames<RangeFunctionKind>())
            Assert.Contains($"RangeFunction(,A1:A20,{kind})", service.Calls);
    }

    [Fact]
    public void A_range_function_reports_what_the_workbook_returns()
    {
        var service = new FakeExcelService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new SumRange
            {
                CellRange = new InArgument<string>("A1:A20"),
                Output = run.Capture<double>("Output"),
            },
            service);

        Assert.Equal(42d, outputs.Values["Output"]);
    }

    [Fact]
    public void A_range_function_rejects_a_missing_range()
    {
        var service = new FakeExcelService();

        Assert.ThrowsAny<Exception>(() => Harness.Run(new SumRange(), service));
    }

    [Fact]
    public void A_chart_activity_prefers_the_name_over_the_index()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new ChartDelete
            {
                SheetName = new InArgument<string>("Data"),
                ChartName = new InArgument<string>("Revenue"),
                ChartIndex = new InArgument<int>(2),
            },
            service);

        Assert.Equal("Revenue", service.LastChart!.ChartName);
        Assert.Contains("ChartDelete@Data/Revenue", service.Calls);
    }

    [Fact]
    public void ChartFormat_passes_zero_measurements_through_untouched()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new ChartFormat
            {
                ChartIndex = new InArgument<int>(1),
                ChartWidth = new InArgument<double>(320),
            },
            service);

        var bounds = Assert.IsType<ChartBounds>(service.LastBounds);
        Assert.Equal(320, bounds.Width);
        // Left unset, so the service is told to leave those measurements alone.
        Assert.Equal(0, bounds.Left);
        Assert.Equal(0, bounds.Height);
    }

    [Fact]
    public void ChartEmbedToPowerPoint_carries_its_target_slide()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new ChartEmbedToPowerPoint
            {
                ChartIndex = new InArgument<int>(1),
                PptFile = new InArgument<string>(@"C:\decks\deck.pptx"),
                SlideIndex = new InArgument<int>(4),
            },
            service);

        Assert.Contains(@"ChartEmbedToPowerPoint(C:\decks\deck.pptx,slide=4)@/chart1", service.Calls);
    }

    [Fact]
    public void The_empty_row_activities_report_the_rows_they_acted_on()
    {
        var service = new FakeExcelService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new DeleteEmptyRows
            {
                StartRowIndex = new InArgument<long>(2),
                AffectedRows = run.Capture<long[]>("AffectedRows"),
            },
            service);

        Assert.Equal(new long[] { 4, 9 }, Assert.IsType<long[]>(outputs.Values["AffectedRows"]));
        Assert.Contains("DeleteEmptyRows(,2)", service.Calls);
    }

    [Fact]
    public void HideUnhideEmptyRows_forwards_its_hide_or_show_choice()
    {
        var service = new FakeExcelService();

        Harness.Run(new HideUnhideEmptyRows { HideDelete = HideDeleteEnum.UnHide }, service);

        Assert.Contains("HideUnhideEmptyRows(,1,UnHide)", service.Calls);
    }

    [Fact]
    public void The_freeze_activities_forward_their_counts_and_their_choice()
    {
        var service = new FakeExcelService();

        Harness.Run(new FreezeRows { NoRows = new InArgument<int>(2) }, service);
        Harness.Run(
            new FreezeColumns
            {
                NoColumns = new InArgument<int>(1),
                FreezeOption = FreezePanesEnum.UnFreeze,
            },
            service);

        Assert.Contains("FreezeRows(,2,Freeze)", service.Calls);
        Assert.Contains("FreezeColumns(,1,UnFreeze)", service.Calls);
    }

    [Fact]
    public void CellFont_gathers_its_arguments_into_one_request()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new CellFont
            {
                CellRange = new InArgument<string>("A1:D1"),
                FontName = new InArgument<string>("Cambria"),
                FontSize = new InArgument<double>(14),
                Font_Style = FontStyleEnum.BoldItalic,
                FontUnderLine = FontUnderLineEnum.Single,
                Strikethrough = SelectionYesNoNone.Yes,
                FontColor = new InArgument<Color>(_ => Color.Red),
            },
            service);

        var font = Assert.IsType<CellFontRequest>(service.LastFont);
        Assert.Equal("Cambria", font.FontName);
        Assert.Equal(14, font.FontSize);
        Assert.Equal(FontStyleEnum.BoldItalic, font.FontStyle);
        Assert.Equal(FontUnderLineEnum.Single, font.FontUnderLine);
        Assert.Equal(SelectionYesNoNone.Yes, font.Strikethrough);
        Assert.Equal(Color.Red, font.FontColor);
        // Left unset, so the service is told to leave the fill alone.
        Assert.True(font.BackgroundColor.IsEmpty);
        // And likewise the script, which has a None of its own.
        Assert.Equal(FontScriptEnum.None, font.FontScript);
    }

    [Fact]
    public void GetRangeStyle_reports_the_settings_it_can()
    {
        var service = new FakeExcelService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new GetRangeStyle
            {
                CellRange = new InArgument<string>("A1"),
                FontName = run.Capture<string>("FontName"),
                FontSize = run.Capture<double>("FontSize"),
                FontColor = run.Capture<Color>("FontColor"),
                BackgroundColor = run.Capture<Color>("BackgroundColor"),
            },
            service);

        Assert.Equal("Calibri", outputs.Values["FontName"]);
        Assert.Equal(11d, outputs.Values["FontSize"]);
        Assert.Equal(Color.FromArgb(255, 0, 0), outputs.Values["FontColor"]);
    }

    [Fact]
    public void ClearSheet_carries_every_switch_it_was_given()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new ClearSheet
            {
                ClearContents = new InArgument<bool>(true),
                ClearHyperlinks = new InArgument<bool>(true),
            },
            service);

        var options = Assert.IsType<ClearOptions>(service.LastClear);
        Assert.True(options.Contents);
        Assert.True(options.Hyperlinks);
        Assert.False(options.All);
        Assert.False(options.Formats);
    }

    [Fact]
    public void Find_reports_both_the_matches_and_the_rows()
    {
        var service = new FakeExcelService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new Find
            {
                FindText = new InArgument<string>("Total"),
                FindOption = FindReplaceEnum.Part,
                MatchCase = true,
                Output = run.Capture<string[]>("Output"),
                RowIndexes = run.Capture<int[]>("RowIndexes"),
            },
            service);

        Assert.Equal(new[] { "Total", "Total" }, Assert.IsType<string[]>(outputs.Values["Output"]));
        Assert.Equal(new[] { 3, 8 }, Assert.IsType<int[]>(outputs.Values["RowIndexes"]));
        Assert.Contains("Find(,,Total,Part,True)", service.Calls);
    }

    [Fact]
    public void FindLastColumn_reports_both_its_number_and_its_letters()
    {
        var service = new FakeExcelService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new FindLastColumn
            {
                ColumnIndex = run.Capture<int>("ColumnIndex"),
                ColumnName = run.Capture<string>("ColumnName"),
            },
            service);

        Assert.Equal(27, outputs.Values["ColumnIndex"]);
        Assert.Equal("AA", outputs.Values["ColumnName"]);
    }

    [Fact]
    public void PageSetup_gathers_all_twenty_one_of_its_arguments()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new PageSetup
            {
                CenterHeader = new InArgument<string>("Quarterly"),
                TopMargin = new InArgument<int>(36),
                PageOrientation = PageOrientationEnum.Landscape,
                PageSize = PaperSizeEnum.PaperA4,
                PrintGridlines = TrueFaleNoneEnum.True,
                ZoomLevel = new InArgument<int>(80),
            },
            service);

        var setup = Assert.IsType<PageSetupRequest>(service.LastPageSetup);
        Assert.Equal("Quarterly", setup.CenterHeader);
        Assert.Equal(36, setup.TopMargin);
        Assert.Equal(PageOrientationEnum.Landscape, setup.PageOrientation);
        Assert.Equal(PaperSizeEnum.PaperA4, setup.PageSize);
        Assert.Equal(TrueFaleNoneEnum.True, setup.PrintGridlines);
        Assert.Equal(80, setup.ZoomLevel);
        // Left unset, so each stays at the value that means "leave it alone".
        Assert.Equal(string.Empty, setup.LeftHeader);
        Assert.Equal(0, setup.BottomMargin);
        Assert.Equal(PageOrderEnum.None, setup.PageOrder);
        Assert.Equal(TrueFaleNoneEnum.None, setup.PrintZoom);
    }

    [Fact]
    public void SetBorder_forwards_the_interop_line_style_it_was_given()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new SetBorder
            {
                CellRange = new InArgument<string>("A1:D10"),
                Presets = BorderEnum.AllBorder,
                LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble,
                BorderWeight = new InArgument<double>(3),
            },
            service);

        Assert.Contains("SetBorder(,A1:D10,AllBorder,xlDouble,3)", service.Calls);
    }

    [Fact]
    public void An_image_activity_prefers_the_name_over_the_index()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new ImageResize
            {
                SheetName = new InArgument<string>("Cover"),
                ImageName = new InArgument<string>("Logo"),
                ImageIndex = new InArgument<int>(3),
                ImageWidth = new InArgument<float>(120),
            },
            service);

        Assert.Equal("Logo", service.LastImage!.ImageName);
        Assert.Contains("ImageResize@Cover/Logo(120x0)", service.Calls);
    }

    [Fact]
    public void ImagesDelete_forwards_both_the_names_and_the_positions()
    {
        var service = new FakeExcelService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new ImagesDelete
            {
                ImageNames = new InArgument<string[]>(_ => new[] { "Logo" }),
                ImageIndexes = new InArgument<int[]>(_ => new[] { 2, 3 }),
                DeletedImageCount = run.Capture<int>("DeletedImageCount"),
            },
            service);

        Assert.Contains("ImagesDelete(,[Logo],[2,3])", service.Calls);
        Assert.Equal(3, outputs.Values["DeletedImageCount"]);
    }

    [Fact]
    public void A_table_activity_prefers_the_name_over_the_index()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new DeleteTable
            {
                TableName = new InArgument<string>("Sales"),
                TableIndex = new InArgument<int>(2),
            },
            service);

        Assert.Equal("Sales", service.LastTable!.TableName);
        Assert.Contains("DeleteTable@Sales", service.Calls);
    }

    [Fact]
    public void GetTableNames_reports_both_shapes_of_its_result()
    {
        var service = new FakeExcelService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new GetTableNames
            {
                Output = run.Capture<List<string>>("Output"),
                OutputTable = run.Capture<DataTable>("OutputTable"),
            },
            service);

        Assert.Equal(new[] { "Sales", "Costs" }, Assert.IsType<List<string>>(outputs.Values["Output"]));
        Assert.Equal("Tables", Assert.IsType<DataTable>(outputs.Values["OutputTable"]).TableName);
    }

    [Fact]
    public void MergeSheetByRow_gathers_its_source_into_one_request()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new MergeSheetByRow
            {
                SheetName = new InArgument<string>("Summary"),
                AppendFileName = new InArgument<string>(@"C:\books\q2.xlsx"),
                AppendSheet = new InArgument<string>("Data"),
                StartColumn = new InArgument<string>("B"),
            },
            service);

        var merge = Assert.IsType<MergeRequest>(service.LastMerge);
        Assert.Equal(@"C:\books\q2.xlsx", merge.AppendFileName);
        Assert.Equal("Data", merge.AppendSheet);
        Assert.Equal("B", merge.StartColumn);
        Assert.Equal("Summary", service.LastSheetName);
    }

    [Fact]
    public void The_add_in_activities_report_what_the_workbook_returns()
    {
        var service = new FakeExcelService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new ExistsAddIns
            {
                AddInsName = new InArgument<string>("Solver"),
                Result = run.Capture<bool>("Result"),
            },
            service);

        Assert.True((bool)outputs.Values["Result"]);

        var listing = new ScopeRun();
        var all = listing.Invoke(
            new GetAllAddins { AddinsList = listing.Capture<Dictionary<string, string>>("AddinsList") },
            new FakeExcelService());

        Assert.Equal(2, Assert.IsType<Dictionary<string, string>>(all.Values["AddinsList"]).Count);
    }

    [Fact]
    public void ListTrustedLocation_reports_both_shapes_of_its_result()
    {
        var service = new FakeExcelService();
        var run = new ScopeRun();

        var outputs = run.Invoke(
            new ListTrustedLocation
            {
                OutputList = run.Capture<string[]>("OutputList"),
                UrlTable = run.Capture<DataTable>("UrlTable"),
            },
            service);

        Assert.Equal(new[] { @"C:\trusted" }, Assert.IsType<string[]>(outputs.Values["OutputList"]));
    }

    [Fact]
    public void The_outline_activities_forward_their_direction()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new GroupRange
            {
                CellRange = new InArgument<string>("3:7"),
                GroupType = GroupEnum.Columns,
            },
            service);

        Assert.Contains("GroupRange(,3:7,Columns)", service.Calls);
    }

    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(52, "AZ")]
    [InlineData(53, "BA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    public void GetColumnName_walks_Excels_bijective_base_twenty_six(int index, string expected)
    {
        var captured = new Outputs();
        var name = new Variable<string>();
        var root = new System.Activities.Statements.Sequence
        {
            Variables = { name },
            Activities =
            {
                new GetColumnName
                {
                    ColumnIndex = new InArgument<int>(index),
                    ColumnName = new OutArgument<string>(name),
                },
                new CaptureOutput<string>(captured, "ColumnName")
                {
                    Value = new InArgument<string>(name),
                },
            },
        };

        new WorkflowInvoker(root).Invoke();
        Assert.Equal(expected, captured.Values["ColumnName"]);
    }

    [Theory]
    [InlineData("A", 1)]
    [InlineData("Z", 26)]
    [InlineData("aa", 27)]
    [InlineData("ZZ", 702)]
    [InlineData("AAA", 703)]
    public void GetColumnNumber_is_the_other_half_of_that(string name, int expected)
    {
        var captured = new Outputs();
        var index = new Variable<int>();
        var root = new System.Activities.Statements.Sequence
        {
            Variables = { index },
            Activities =
            {
                new GetColumnNumber
                {
                    ColumnName = new InArgument<string>(name),
                    ColumnIndex = new OutArgument<int>(index),
                },
                new CaptureOutput<int>(captured, "ColumnIndex")
                {
                    Value = new InArgument<int>(index),
                },
            },
        };

        new WorkflowInvoker(root).Invoke();
        Assert.Equal(expected, captured.Values["ColumnIndex"]);
    }

    [Theory]
    [InlineData("")]
    [InlineData("A1")]
    [InlineData("A-B")]
    public void GetColumnNumber_rejects_what_is_not_a_column_name(string name)
    {
        var activity = new GetColumnNumber
        {
            ColumnName = new InArgument<string>(name),
            ColumnIndex = new OutArgument<int>(new Variable<int>()),
        };

        Assert.ThrowsAny<Exception>(() => new WorkflowInvoker(activity).Invoke());
    }

    [Fact]
    public void GetColumnName_rejects_a_column_before_the_first()
    {
        var activity = new GetColumnName
        {
            ColumnIndex = new InArgument<int>(0),
            ColumnName = new OutArgument<string>(new Variable<string>()),
        };

        Assert.ThrowsAny<Exception>(() => new WorkflowInvoker(activity).Invoke());
    }

    [Fact]
    public void SaveAsSheet_stands_outside_a_scope_and_gathers_its_own_request()
    {
        var service = new FakeExcelService();

        Harness.RunAlone(
            new SaveAsSheet
            {
                FileName = new InArgument<string>(@"C:\books\book.xlsx"),
                NewFileName = new InArgument<string>(@"C:\books\sheet.xlsx"),
                Sheet = new InArgument<string>("Summary"),
            },
            service);

        var request = Assert.IsType<SaveSheetRequest>(service.LastSaveSheet);
        Assert.Equal(@"C:\books\book.xlsx", request.FileName);
        Assert.Equal(@"C:\books\sheet.xlsx", request.NewFileName);
        Assert.Equal("Summary", request.Sheet);
    }

    [Fact]
    public void SaveAsSheet_rejects_a_missing_destination()
    {
        var service = new FakeExcelService();
        var activity = new SaveAsSheet { FileName = new InArgument<string>(@"C:\books\book.xlsx") };

        Assert.ThrowsAny<Exception>(() => Harness.RunAlone(activity, service));
    }

    [Fact]
    public void CloseAllExcel_stands_outside_a_scope_too()
    {
        var service = new FakeExcelService();

        Harness.RunAlone(new CloseAllExcel(), service);

        Assert.Contains("CloseAllExcel()", service.Calls);
    }

    private static Exception Unwrap(Exception error) =>
        error.InnerException is null ? error : Unwrap(error.InnerException);
}
