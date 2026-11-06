using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

public class Program
{
    // Path to the static logo image that will be inserted for every merge.
    private const string LogoFileName = "logo.png";

    public static void Main()
    {
        // Ensure the static logo image exists.
        CreateSampleLogo();

        // Create a template document with merge fields.
        string templatePath = "Template.docx";
        CreateTemplateDocument(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Set the callback that will handle image field merging.
        doc.MailMerge.FieldMergingCallback = new ImageFieldMergingHandler(LogoFileName);

        // Prepare simple data source.
        string[] fieldNames = { "Name", "CompanyLogo" };
        object[] fieldValues = { "John Doe", null }; // Value for image field is ignored.

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        string outputPath = "Merged.docx";
        doc.Save(outputPath);

        // Simple validation: ensure the output file exists and is not empty.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
            throw new InvalidOperationException("Merged document was not created correctly.");
    }

    // Creates a tiny PNG logo file from a base64 string.
    private static void CreateSampleLogo()
    {
        // 1x1 pixel transparent PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAn8B9p6XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(LogoFileName, imageBytes);
    }

    // Builds a simple template containing a text field and an image field.
    private static void CreateTemplateDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Employee Report");
        builder.Writeln();

        // Insert a merge field for the employee name.
        builder.InsertField("MERGEFIELD Name");
        builder.Writeln();

        // Insert a merge field that will be replaced by the company logo.
        builder.InsertField("MERGEFIELD CompanyLogo");
        builder.Writeln();

        doc.Save(path);
    }

    // Callback class that handles image field merging.
    private class ImageFieldMergingHandler : IFieldMergingCallback
    {
        private readonly string _staticImagePath;

        public ImageFieldMergingHandler(string staticImagePath)
        {
            _staticImagePath = staticImagePath;
        }

        // Called for regular (non-image) fields; not used here.
        public void FieldMerging(FieldMergingArgs args)
        {
            // No custom handling required for text fields.
        }

        // Called when an image merge field is encountered.
        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Insert the static logo image regardless of the source data.
            args.ImageFileName = _staticImagePath;

            // Set image dimensions (in points) using MergeFieldImageDimension.
            args.ImageWidth = new MergeFieldImageDimension(100);   // Width in points.
            args.ImageHeight = new MergeFieldImageDimension(50);   // Height in points.
        }
    }
}
