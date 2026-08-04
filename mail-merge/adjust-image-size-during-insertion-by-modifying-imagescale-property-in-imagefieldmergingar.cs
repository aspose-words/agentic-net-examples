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
        // Prepare a temporary folder for the demo files.
        string demoDir = Path.Combine(Path.GetTempPath(), "AsposeMailMergeDemo");
        Directory.CreateDirectory(demoDir);

        // Create a simple 1x1 PNG image using a hard‑coded byte array (avoids System.Drawing).
        string imagePath = Path.Combine(demoDir, "SampleImage.png");
        byte[] pngData = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BFwAE/wJ/lK5XAAAAAElFTkSuQmCC");
        File.WriteAllBytes(imagePath, pngData);

        // Create a new document and insert an image merge field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // The field name must start with "Image:" to be recognized as an image merge field.
        builder.InsertField("MERGEFIELD Image:Photo");

        // Build a data source containing the path to the image file.
        DataTable data = new DataTable("Images");
        data.Columns.Add("Photo", typeof(string));
        data.Rows.Add(imagePath);

        // Set up a callback that will adjust the image size during the merge.
        doc.MailMerge.FieldMergingCallback = new ImageResizer(100, 100, MergeFieldImageDimensionUnit.Point);

        // Execute the mail merge.
        doc.MailMerge.Execute(data);
        doc.UpdateFields();

        // Save the resulting document.
        string outputPath = Path.Combine(demoDir, "MergedResult.docx");
        doc.Save(outputPath);
    }

    // Callback that sets the image file name and overrides its dimensions.
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
            // No custom processing required for non‑image fields.
        }

        // Called for each image merge field.
        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Provide the image file name from the data source.
            args.ImageFileName = args.FieldValue.ToString();

            // Override the image dimensions.
            args.ImageWidth = new MergeFieldImageDimension(_width, _unit);
            args.ImageHeight = new MergeFieldImageDimension(_height, _unit);
        }
    }
}
