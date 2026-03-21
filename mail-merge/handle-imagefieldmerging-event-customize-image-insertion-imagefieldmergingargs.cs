using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

class CustomImageHandler : IFieldMergingCallback
{
    void IFieldMergingCallback.FieldMerging(FieldMergingArgs args) { }

    void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
    {
        string key = args.FieldValue?.ToString();
        if (string.IsNullOrEmpty(key))
            return;

        if (ImageMap.TryGetValue(key, out string filePath) && File.Exists(filePath))
        {
            // Provide the image file name; Aspose.Words will load the image from the file.
            args.ImageFileName = filePath;
        }
    }

    private static readonly Dictionary<string, string> ImageMap = new()
    {
        { "Logo", Path.Combine(Path.GetTempPath(), "Logo.png") },
        { "Badge", Path.Combine(Path.GetTempPath(), "Badge.png") }
    };
}

class Program
{
    static void Main()
    {
        // Ensure temporary images exist.
        CreateSampleImage(Path.Combine(Path.GetTempPath(), "Logo.png"));
        CreateSampleImage(Path.Combine(Path.GetTempPath(), "Badge.png"));

        // Create a simple template document with an image merge field and a text merge field.
        string templatePath = Path.Combine(Path.GetTempPath(), "ReportTemplate.docx");
        CreateTemplateDocument(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Attach the custom callback that supplies images.
        doc.MailMerge.FieldMergingCallback = new CustomImageHandler();

        // Build a data source. The "Photo" column holds the short identifiers used above.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name");
        table.Columns.Add("Photo");
        table.Rows.Add("John Doe", "Logo");
        table.Rows.Add("Jane Smith", "Badge");

        // Perform the mail merge.
        doc.MailMerge.Execute(table);

        // Save the merged document.
        string outputPath = Path.Combine(Path.GetTempPath(), "ReportMerged.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Merged document saved to: {outputPath}");
    }

    private static void CreateSampleImage(string path)
    {
        // A 1x1 pixel transparent PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, pngBytes);
    }

    private static void CreateTemplateDocument(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert an image merge field.
        builder.InsertField("MERGEFIELD Image:Photo \\* MERGEFORMAT");
        builder.Writeln();

        // Insert a text merge field.
        builder.InsertField("MERGEFIELD Name \\* MERGEFORMAT");
        builder.Writeln();

        doc.Save(path);
    }
}
