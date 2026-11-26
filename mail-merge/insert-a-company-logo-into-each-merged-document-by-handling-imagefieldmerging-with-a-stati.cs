using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Prepare a simple logo image file (1x1 red pixel PNG).
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "logo.png");
        CreateSampleLogo(imagePath);

        // Create a template document with merge fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Company Name: ");
        builder.InsertField("MERGEFIELD CompanyName \\* MERGEFORMAT");
        builder.Writeln();
        builder.Writeln("Company Logo:");
        builder.InsertField("MERGEFIELD CompanyLogo \\* MERGEFORMAT");
        builder.Writeln();

        // Set up the image field merging callback.
        doc.MailMerge.FieldMergingCallback = new ImageFieldMergingCallback(imagePath);

        // Prepare data for the merge.
        string[] fieldNames = { "CompanyName", "CompanyLogo" };
        object[] fieldValues = { "Acme Corp", null }; // Image will be supplied by the callback.

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Validate that the text field was merged.
        if (!doc.GetText().Contains("Acme Corp"))
            throw new InvalidOperationException("Text merge field was not replaced correctly.");

        // Save the merged document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocument.docx");
        doc.Save(outputPath);

        // Validation: ensure the output file exists and is not empty.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
            throw new InvalidOperationException("Merged document was not created correctly.");

        // Clean up the temporary logo file.
        if (File.Exists(imagePath))
            File.Delete(imagePath);
    }

    private static void CreateSampleLogo(string path)
    {
        // Base64 for a 1x1 red pixel PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lKcVAAAAAElFTkSuQmCC";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, pngBytes);
    }

    private class ImageFieldMergingCallback : IFieldMergingCallback
    {
        private readonly string _imagePath;

        public ImageFieldMergingCallback(string imagePath)
        {
            _imagePath = imagePath ?? throw new ArgumentNullException(nameof(imagePath));
        }

        public void FieldMerging(FieldMergingArgs args)
        {
            // No custom handling for text fields.
        }

        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            if (args.FieldName == "CompanyLogo" && File.Exists(_imagePath))
            {
                // Provide the image via a stream to avoid System.Drawing dependencies.
                args.ImageStream = File.OpenRead(_imagePath);
            }
        }
    }
}
