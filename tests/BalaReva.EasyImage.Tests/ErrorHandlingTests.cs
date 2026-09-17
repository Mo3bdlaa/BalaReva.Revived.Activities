using System.Activities;

namespace BalaReva.EasyImage.Tests;

public class ErrorHandlingTests
{
    [Fact]
    public void A_missing_source_file_names_the_file()
    {
        using var images = new TestImages();
        var activity = new ImageSizes
        {
            ImagePath = new InArgument<string>(images.At("absent.png")),
        };

        var error = Assert.ThrowsAny<Exception>(() => TestImages.Run(activity));
        Assert.Contains("absent.png", Unwrap(error).Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ContinueOnError_reports_failure_instead_of_faulting()
    {
        using var images = new TestImages();
        var activity = new ImageSizes
        {
            ContinueOnError = new InArgument<bool>(true),
            ImagePath = new InArgument<string>(images.At("absent.png")),
        };

        var outputs = TestImages.Run(activity);

        Assert.False((bool)outputs["ExecutionResult"]);
    }

    [Fact]
    public void A_successful_activity_reports_ExecutionResult_true()
    {
        using var images = new TestImages();
        var activity = new ImageSizes
        {
            ImagePath = new InArgument<string>(images.Create("in.png", 10, 10)),
        };

        var outputs = TestImages.Run(activity);

        Assert.True((bool)outputs["ExecutionResult"]);
    }

    [Fact]
    public void A_blank_destination_path_is_rejected()
    {
        using var images = new TestImages();
        var activity = new ImageResize
        {
            ImagePath = new InArgument<string>(images.Create("in.png", 10, 10)),
            DesImagePath = new InArgument<string>("  "),
            ImageWidth = new InArgument<int>(5),
            ImageHeight = new InArgument<int>(5),
        };

        Assert.ThrowsAny<Exception>(() => TestImages.Run(activity));
    }

    [Fact]
    public void A_destination_folder_that_does_not_exist_yet_is_created()
    {
        using var images = new TestImages();
        var destination = Path.Combine(images.Folder, "nested", "deeper", "out.png");

        TestImages.Run(new ImageResize
        {
            ImagePath = new InArgument<string>(images.Create("in.png", 20, 20)),
            DesImagePath = new InArgument<string>(destination),
            ImageWidth = new InArgument<int>(10),
            ImageHeight = new InArgument<int>(10),
        });

        Assert.True(File.Exists(destination));
    }

    private static Exception Unwrap(Exception error) =>
        error.InnerException is null ? error : Unwrap(error.InnerException);
}
