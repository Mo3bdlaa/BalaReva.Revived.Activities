using BalaReva.EasyExcel.Utilities;
using Interop = Microsoft.Office.Interop.Excel;

namespace BalaReva.EasyExcel;

public sealed partial class ExcelWorkbook
{
    /// <summary>The picture on a sheet, by name when given and by position otherwise.</summary>
    private Interop.Shape? PictureOn(ImageRef image)
    {
        ArgumentNullException.ThrowIfNull(image);
        var pictures = Pictures(Sheet(image.SheetName));

        if (!string.IsNullOrWhiteSpace(image.ImageName))
            return pictures.FirstOrDefault(shape =>
                string.Equals(shape.Name, image.ImageName, StringComparison.OrdinalIgnoreCase));

        var index = image.ImageIndex < 1 ? 1 : image.ImageIndex;
        return index <= pictures.Count ? pictures[index - 1] : null;
    }

    /// <summary>Every picture shape on a sheet, charts and other shapes left out.</summary>
    private static List<Interop.Shape> Pictures(Interop.Worksheet sheet)
    {
        var pictures = new List<Interop.Shape>();
        foreach (Interop.Shape shape in sheet.Shapes)
        {
            // msoPicture and msoLinkedPicture, as numbers: Shape.Type is an
            // MsoShapeType, which lives in office.dll. See Mso in ExcelService.cs.
            int type = ((dynamic)shape).Type;
            if (type is Mso.Picture or Mso.LinkedPicture) pictures.Add(shape);
        }
        return pictures;
    }

    /// <inheritdoc />
    public bool ImageExists(ImageRef image) => PictureOn(image) is not null;

    /// <inheritdoc />
    public bool ImageResize(ImageRef image, float width, float height)
    {
        var picture = PictureOn(image);
        if (picture is null) return false;

        // Excel keeps the aspect ratio locked by default, which would quietly undo one of
        // the two measurements. LockAspectRatio is an MsoTriState, so it is set as a
        // number through dynamic.
        ((dynamic)picture).LockAspectRatio = Mso.False;

        if (width > 0) picture.Width = width;
        if (height > 0) picture.Height = height;

        Save();
        return true;
    }

    /// <inheritdoc />
    public int ImagesDelete(string sheetName, string[] imageNames, int[] imageIndexes)
    {
        var pictures = Pictures(Sheet(sheetName));
        var names = new HashSet<string>(imageNames ?? [], StringComparer.OrdinalIgnoreCase);
        var indexes = new HashSet<int>(imageIndexes ?? []);

        var doomed = pictures
            .Where((shape, position) => names.Contains(shape.Name) || indexes.Contains(position + 1))
            .ToList();

        foreach (var shape in doomed) shape.Delete();
        if (doomed.Count > 0) Save();
        return doomed.Count;
    }

    /// <inheritdoc />
    public int ImagesDeleteAll(string sheetName)
    {
        var pictures = Pictures(Sheet(sheetName));
        foreach (var shape in pictures) shape.Delete();
        if (pictures.Count > 0) Save();
        return pictures.Count;
    }

    /// <inheritdoc />
    public void CopyAsPicture(string sheetName, string cell, string imageFilePath)
    {
        var sheet = Sheet(sheetName);
        var range = Target(sheet, cell);

        range.CopyPicture(
            Interop.XlPictureAppearance.xlScreen,
            Interop.XlCopyPictureFormat.xlBitmap);

        // Excel can put a picture on the clipboard but cannot write one to a file. A
        // chart sheet sized to the range is the usual way round it: paste into it and
        // export, which Excel can do.
        var chart = (Interop.Chart)_workbook.Charts.Add();
        try
        {
            chart.ChartArea.Width = Convert.ToDouble(range.Width);
            chart.ChartArea.Height = Convert.ToDouble(range.Height);
            chart.Paste();
            chart.Export(imageFilePath, Path.GetExtension(imageFilePath).TrimStart('.').ToUpperInvariant());
        }
        finally
        {
            var alerts = _application.DisplayAlerts;
            _application.DisplayAlerts = false;
            chart.Delete();
            _application.DisplayAlerts = alerts;
        }
    }

    /// <inheritdoc />
    public void ImageExtractor(string sheetName, string imageFolder, FileExtension extension)
    {
        Directory.CreateDirectory(imageFolder);
        var sheet = Sheet(sheetName);
        var suffix = extension.ToString().ToLowerInvariant();
        var index = 0;

        foreach (var picture in Pictures(sheet))
        {
            index++;
            var name = string.IsNullOrWhiteSpace(picture.Name) ? $"Image{index}" : picture.Name;
            picture.Copy();

            var chart = (Interop.Chart)_workbook.Charts.Add();
            try
            {
                chart.ChartArea.Width = picture.Width;
                chart.ChartArea.Height = picture.Height;
                chart.Paste();
                chart.Export(Path.Combine(imageFolder, $"{name}.{suffix}"), suffix.ToUpperInvariant());
            }
            finally
            {
                var alerts = _application.DisplayAlerts;
                _application.DisplayAlerts = false;
                chart.Delete();
                _application.DisplayAlerts = alerts;
            }
        }
    }
}
