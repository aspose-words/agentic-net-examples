using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a temporary folder for sample images.
        string tempImageFolder = Path.Combine(Path.GetTempPath(), "AsposeSampleImages");
        Directory.CreateDirectory(tempImageFolder);

        // Create two tiny PNG images (1x1 pixel) from a base‑64 string.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        string imagePath1 = Path.Combine(tempImageFolder, "Image1.png");
        string imagePath2 = Path.Combine(tempImageFolder, "Image2.png");
        File.WriteAllBytes(imagePath1, pngBytes);
        File.WriteAllBytes(imagePath2, pngBytes);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an image merge field that will receive images during mail merge.
        // The field name must start with "Image:" to be recognized as an image field.
        builder.InsertField("MERGEFIELD Image:Photo");

        // Prepare a simple data source containing file paths to the images.
        DataTable table = new DataTable("Images");
        table.Columns.Add("Photo");
        table.Rows.Add(imagePath1);
        table.Rows.Add(imagePath2);

        // Attach a callback that will resize every merged image to the desired dimensions.
        doc.MailMerge.FieldMergingCallback = new ImageResizer(150, 150, MergeFieldImageDimensionUnit.Point);

        // Perform the mail merge.
        doc.MailMerge.Execute(table);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ResizedImages.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");

        // Clean up temporary images.
        try { File.Delete(imagePath1); } catch { }
        try { File.Delete(imagePath2); } catch { }
        try { Directory.Delete(tempImageFolder, true); } catch { }
    }

    // Implementation of IFieldMergingCallback that resizes images.
    private class ImageResizer : IFieldMergingCallback
    {
        private readonly double _width;
        private readonly double _height;
        private readonly MergeFieldImageDimensionUnit _unit;

        public ImageResizer(double width, double height, MergeFieldImageDimensionUnit unit)
        {
            _width = width;
            _height = height;
            _unit = unit;
        }

        // Not used for plain text fields.
        public void FieldMerging(FieldMergingArgs args) { }

        // Called for each image merge field.
        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Supply the image file name from the data source.
            args.ImageFileName = args.FieldValue?.ToString();

            // Override the default dimensions with the desired size.
            args.ImageWidth = new MergeFieldImageDimension(_width, _unit);
            args.ImageHeight = new MergeFieldImageDimension(_height, _unit);
        }
    }
}
