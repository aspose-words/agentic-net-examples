using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a temporary PNG image file without using System.Drawing.
        // This is a 1x1 pixel transparent PNG encoded in base64.
        string tempImagePath = Path.Combine(Path.GetTempPath(), "TempImage.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(tempImagePath, pngBytes);

        // Create a new document and insert an image merge field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // The field name must start with "Image:" so the mail merge engine treats it as an image field.
        builder.InsertField("MERGEFIELD Image:Photo");

        // Prepare a data source with a single column that contains the image file name.
        DataTable table = new DataTable("Images");
        table.Columns.Add("Photo", typeof(string));
        table.Rows.Add(tempImagePath);

        // Set a callback that will resize the image during the merge.
        doc.MailMerge.FieldMergingCallback = new ImageResizer(150, 150, MergeFieldImageDimensionUnit.Point);
        doc.MailMerge.Execute(table);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedImage.docx");
        doc.Save(outputPath);

        // Clean up the temporary image file.
        if (File.Exists(tempImagePath))
            File.Delete(tempImagePath);
    }

    // Callback that sets the image size for each merged image.
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

        // Not used for text fields.
        public void FieldMerging(FieldMergingArgs args)
        {
            // No action required.
        }

        // Called for image merge fields.
        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Provide the image file name from the data source.
            args.ImageFileName = args.FieldValue.ToString();

            // Override the default dimensions.
            args.ImageWidth = new MergeFieldImageDimension(_width, _unit);
            args.ImageHeight = new MergeFieldImageDimension(_height, _unit);
        }
    }
}
