using System.Activities;
using System.Drawing;
using System.Drawing.Imaging;

namespace BalaReva.EasyImage.Tests;

/// <summary>
/// Each activity runs as a real workflow against real image files. These need GDI+,
/// so the suite runs on the Windows CI job.
/// </summary>
public class ActivityTests
{
    [Fact]
    public void ImageSizes_reports_the_pixel_dimensions()
    {
        using var images = new TestImages();
        var source = images.Create("in.png", 120, 80);

        var outputs = TestImages.Run(new ImageSizes { ImagePath = new InArgument<string>(source) });

        Assert.Equal(120, outputs["ImageWidth"]);
        Assert.Equal(80, outputs["ImageHeight"]);
    }

    [Fact]
    public void ImageResize_writes_an_image_of_the_requested_size()
    {
        using var images = new TestImages();
        var source = images.Create("in.png", 120, 80);
        var destination = images.At("out.png");

        TestImages.Run(new ImageResize
        {
            ImagePath = new InArgument<string>(source),
            DesImagePath = new InArgument<string>(destination),
            ImageWidth = new InArgument<int>(60),
            ImageHeight = new InArgument<int>(40),
        });

        Assert.Equal((60, 40), TestImages.SizeOf(destination));
    }

    [Fact]
    public void ImageResize_can_overwrite_its_own_source()
    {
        // Image.FromFile would hold a lock on the source and make this fail.
        using var images = new TestImages();
        var source = images.Create("in.png", 120, 80);

        TestImages.Run(new ImageResize
        {
            ImagePath = new InArgument<string>(source),
            DesImagePath = new InArgument<string>(source),
            ImageWidth = new InArgument<int>(30),
            ImageHeight = new InArgument<int>(20),
        });

        Assert.Equal((30, 20), TestImages.SizeOf(source));
    }

    [Fact]
    public void ImageCropper_writes_just_the_requested_rectangle()
    {
        using var images = new TestImages();
        var source = images.Create("in.png", 120, 80);
        var destination = images.At("out.png");

        TestImages.Run(new ImageCropper
        {
            ImagePath = new InArgument<string>(source),
            DesImagePath = new InArgument<string>(destination),
            ImgX = new InArgument<int>(10),
            ImgY = new InArgument<int>(20),
            ImgWidth = new InArgument<int>(50),
            ImgHeight = new InArgument<int>(30),
        });

        Assert.Equal((50, 30), TestImages.SizeOf(destination));
    }

    [Fact]
    public void ImageCropper_rejects_a_rectangle_that_runs_off_the_image()
    {
        // GDI+ would quietly clamp this and produce a smaller image than asked for.
        using var images = new TestImages();
        var source = images.Create("in.png", 100, 100);

        var activity = new ImageCropper
        {
            ImagePath = new InArgument<string>(source),
            DesImagePath = new InArgument<string>(images.At("out.png")),
            ImgX = new InArgument<int>(80),
            ImgY = new InArgument<int>(0),
            ImgWidth = new InArgument<int>(50),
            ImgHeight = new InArgument<int>(50),
        };

        Assert.ThrowsAny<Exception>(() => TestImages.Run(activity));
        Assert.False(File.Exists(images.At("out.png")));
    }

    [Fact]
    public void ImageRotate_swaps_the_axes_on_a_quarter_turn()
    {
        using var images = new TestImages();
        var source = images.Create("in.png", 120, 80);
        var destination = images.At("out.png");

        TestImages.Run(new ImageRotate
        {
            ImagePath = new InArgument<string>(source),
            DesImagePath = new InArgument<string>(destination),
            FlipType = RotateFlipType.Rotate90FlipNone,
        });

        Assert.Equal((80, 120), TestImages.SizeOf(destination));
    }

    [Fact]
    public void ImageConverter_writes_the_format_that_was_asked_for()
    {
        using var images = new TestImages();
        var source = images.Create("in.png", 40, 40);
        var destination = images.At("out.jpg");

        TestImages.Run(new ImageConverter
        {
            ImagePath = new InArgument<string>(source),
            DesImagePath = new InArgument<string>(destination),
            FileFormat = EnumFileFormat.JPEG,
        });

        Assert.Equal(ImageFormat.Jpeg.Guid, TestImages.FormatOf(destination).Guid);
    }

    [Fact]
    public void ImageCompression_at_a_lower_quality_makes_a_smaller_file()
    {
        using var images = new TestImages();
        // A flat colour compresses to almost nothing at any quality, so use noise.
        var source = images.At("in.png");
        using (var bitmap = new Bitmap(200, 200))
        {
            var random = new Random(1);
            for (var x = 0; x < 200; x++)
            for (var y = 0; y < 200; y++)
                bitmap.SetPixel(x, y, Color.FromArgb(random.Next(256), random.Next(256), random.Next(256)));
            bitmap.Save(source, ImageFormat.Png);
        }

        var low = images.At("low.jpg");
        var high = images.At("high.jpg");
        TestImages.Run(new ImageCompression
        {
            ImagePath = new InArgument<string>(source),
            DesImagePath = new InArgument<string>(low),
            ImageQuality = new InArgument<int>(10),
        });
        TestImages.Run(new ImageCompression
        {
            ImagePath = new InArgument<string>(source),
            DesImagePath = new InArgument<string>(high),
            ImageQuality = new InArgument<int>(95),
        });

        Assert.True(new FileInfo(low).Length < new FileInfo(high).Length,
            "Quality 10 should produce a smaller file than quality 95.");
    }

    [Fact]
    public void ImageCompression_rejects_a_quality_outside_0_to_100()
    {
        using var images = new TestImages();
        var activity = new ImageCompression
        {
            ImagePath = new InArgument<string>(images.Create("in.png", 10, 10)),
            DesImagePath = new InArgument<string>(images.At("out.jpg")),
            ImageQuality = new InArgument<int>(150),
        };

        Assert.ThrowsAny<Exception>(() => TestImages.Run(activity));
    }

    [Fact]
    public void ImageCombine_lays_the_two_images_out_side_by_side()
    {
        using var images = new TestImages();
        var first = images.Create("a.png", 60, 40);
        var second = images.Create("b.png", 30, 50, Color.Firebrick);
        var destination = images.At("out.png");

        TestImages.Run(new ImageCombine
        {
            ImagePath1 = new InArgument<string>(first),
            ImagePath2 = new InArgument<string>(second),
            DesImagePath = new InArgument<string>(destination),
        });

        // Widths add, height takes the taller of the two.
        Assert.Equal((90, 50), TestImages.SizeOf(destination));
    }

    [Fact]
    public void ImageMerge_stacks_the_two_images()
    {
        using var images = new TestImages();
        var first = images.Create("a.png", 60, 40);
        var second = images.Create("b.png", 30, 50, Color.Firebrick);
        var destination = images.At("out.png");

        TestImages.Run(new ImageMerge
        {
            ImagePath1 = new InArgument<string>(first),
            ImagePath2 = new InArgument<string>(second),
            DesImagePath = new InArgument<string>(destination),
        });

        // Heights add, width takes the wider of the two.
        Assert.Equal((60, 90), TestImages.SizeOf(destination));
    }

    [Fact]
    public void ImageWatermark_writes_into_the_destination_folder_and_reports_the_path()
    {
        using var images = new TestImages();
        var source = images.Create("photo.png", 200, 100);
        var destination = Path.Combine(images.Folder, "out");

        var outputs = TestImages.Run(new ImageWatermark
        {
            SourceFilePath = new InArgument<string>(source),
            DestinationFolder = new InArgument<string>(destination),
            Text = new InArgument<string>("CONFIDENTIAL"),
            TextPosition = new InArgument<Point>(new Point(5, 5)),
            ForeColor = new InArgument<Color>(Color.Red),
        });

        var written = Assert.IsType<string>(outputs["OutputFile"]);
        Assert.True(File.Exists(written));
        Assert.Equal(destination, Path.GetDirectoryName(written));
        // The watermark is drawn onto the image; the dimensions do not change.
        Assert.Equal((200, 100), TestImages.SizeOf(written));
    }

    [Fact]
    public void ImageWatermark_honours_the_requested_format()
    {
        using var images = new TestImages();
        var source = images.Create("photo.png", 60, 60);

        var outputs = TestImages.Run(new ImageWatermark
        {
            SourceFilePath = new InArgument<string>(source),
            DestinationFolder = new InArgument<string>(images.Folder),
            Text = new InArgument<string>("X"),
            // Not new InArgument<ImageFormat>(ImageFormat.Jpeg): that builds a
            // Literal<ImageFormat>, and WF literals only accept value types and string.
            // A real workflow has to bind this as an expression too. Same for Font.
            ImageFormat = new InArgument<ImageFormat>(_ => ImageFormat.Jpeg),
        });

        var written = (string)outputs["OutputFile"];
        Assert.EndsWith(".jpg", written, StringComparison.Ordinal);
        Assert.Equal(ImageFormat.Jpeg.Guid, TestImages.FormatOf(written).Guid);
    }

    [Fact]
    public void ImageWatermark_accepts_a_font_supplied_by_the_workflow()
    {
        using var images = new TestImages();
        var source = images.Create("photo.png", 300, 120);
        using var font = new Font("Arial", 18, FontStyle.Italic);

        var outputs = TestImages.Run(new ImageWatermark
        {
            SourceFilePath = new InArgument<string>(source),
            DestinationFolder = new InArgument<string>(images.Folder),
            Text = new InArgument<string>("DRAFT"),
            Font = new InArgument<Font>(_ => font),
        });

        Assert.True(File.Exists((string)outputs["OutputFile"]));
        // The activity must not dispose a font it did not create; this would throw if it had.
        Assert.Equal(18, font.Size);
    }
}
