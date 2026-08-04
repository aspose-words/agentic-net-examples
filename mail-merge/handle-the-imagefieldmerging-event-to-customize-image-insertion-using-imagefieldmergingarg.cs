using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary directory for generated files.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeMailMergeDemo");
        Directory.CreateDirectory(tempDir);

        // Create two simple PNG images (1x1 pixel) and save them to disk.
        string redImagePath = Path.Combine(tempDir, "Red.png");
        string greenImagePath = Path.Combine(tempDir, "Green.png");
        WritePngFromBase64(redImagePath, RedPngBase64);
        WritePngFromBase64(greenImagePath, GreenPngBase64);

        // Build a data source that contains short names referencing the images.
        DataTable dataTable = new DataTable("Images");
        dataTable.Columns.Add("ImageColumn", typeof(string));
        dataTable.Rows.Add("Red");
        dataTable.Rows.Add("Green");

        // Map the short names to the actual file paths.
        var imageMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Red", redImagePath },
            { "Green", greenImagePath }
        };

        // Create a new document and insert an image merge field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // The field name includes the "Image:" prefix so the mail merge engine knows it is an image field.
        builder.InsertField("MERGEFIELD Image:ImageColumn");

        // Assign the custom callback that will resolve the short names to actual images.
        doc.MailMerge.FieldMergingCallback = new ImageFilenameCallback(imageMap);

        // Execute the mail merge using the DataTable as the data source.
        doc.MailMerge.Execute(dataTable);

        // Save the resulting document.
        string outputPath = Path.Combine(tempDir, "MergedDocument.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Writes a PNG file from a Base64 string.
    private static void WritePngFromBase64(string filePath, string base64)
    {
        byte[] bytes = Convert.FromBase64String(base64);
        File.WriteAllBytes(filePath, bytes);
    }

    // Base64-encoded 1x1 red PNG.
    private const string RedPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BFwAE/wJ/lKXcAAAAAElFTkSuQmCC";

    // Base64-encoded 1x1 green PNG.
    private const string GreenPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8DAAQAD/AL+K9cVAAAAAElFTkSuQmCC";

    // Custom callback that handles image merge fields.
    private class ImageFilenameCallback : IFieldMergingCallback
    {
        private readonly Dictionary<string, string> _imageFilenames;

        public ImageFilenameCallback(Dictionary<string, string> imageFilenames)
        {
            _imageFilenames = imageFilenames;
        }

        // Not used for text fields in this example.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // No custom processing required.
        }

        // Called when an image merge field is encountered.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            string key = args.FieldValue?.ToString();
            if (!string.IsNullOrEmpty(key) && _imageFilenames.TryGetValue(key, out string fileName))
            {
                // Use the file name directly; no need for System.Drawing.Image.
                args.ImageFileName = fileName;
            }
        }
    }
}
