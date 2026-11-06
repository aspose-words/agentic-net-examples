using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class ImageInfo
{
    // URI can be a file path, a web URL, or a Base64 string.
    public string Uri { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<ImageInfo> Images { get; set; } = new();
}

public class Program
{
    // Returns a valid image URI. If the supplied URI is not a local file that exists,
    // the method falls back to the provided placeholder path.
    private static string ResolveImageUri(string uri, string placeholderPath)
    {
        // Treat HTTP/HTTPS URIs as potentially unreachable for this demo.
        if (uri.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            return placeholderPath;

        // If the path points to an existing file, keep it.
        if (File.Exists(uri))
            return uri;

        // Otherwise use the placeholder.
        return placeholderPath;
    }

    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple PNG image from Base64 data.
        string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAn8B9pXK5WQAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        string validImagePath = Path.Combine(outputDir, "sample.png");
        File.WriteAllBytes(validImagePath, pngBytes);

        // -----------------------------------------------------------------
        // Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin foreach over Images collection.
        builder.Writeln("<<foreach [img in Images]>>");

        // Table that will hold each image.
        Table table = builder.StartTable();
        builder.InsertCell();

        // TextBox container for the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Writeln("<<image [img.Uri] -fitSize>>");

        // Close table row and table.
        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and prepare the data model.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        ReportModel model = new ReportModel();

        // First image is valid.
        model.Images.Add(new ImageInfo
        {
            Uri = ResolveImageUri(validImagePath, validImagePath)
        });

        // Second image URI is invalid; it will be replaced by the placeholder.
        model.Images.Add(new ImageInfo
        {
            Uri = ResolveImageUri("http://nonexistent.invalid/image.png", validImagePath)
        });

        // Configure the ReportingEngine to embed inline error messages.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);

        // Output result status.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}
