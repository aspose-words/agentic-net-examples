using System;
using System.IO;
using System.Net;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Original (potentially unreachable) URL.
    private readonly string _originalUrl = "http://nonexistent.example.com/image.jpg";

    // Returns a reachable image path: the original URL if reachable,
    // otherwise a local placeholder image.
    public string ImagePath
    {
        get
        {
            // Try a simple HEAD request to see if the URL is reachable.
            try
            {
                var request = WebRequest.Create(_originalUrl);
                request.Method = "HEAD";
                using var response = request.GetResponse();
                // If we get a response, assume the image can be loaded.
                return _originalUrl;
            }
            catch
            {
                // Fallback to a local placeholder image.
                return Path.Combine(Directory.GetCurrentDirectory(), "placeholder.png");
            }
        }
    }
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 0. Prepare a local placeholder image (1x1 pixel PNG).
        // -----------------------------------------------------------------
        const string placeholderPath = "placeholder.png";
        if (!File.Exists(placeholderPath))
        {
            // Base64-encoded 1x1 transparent PNG.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(placeholderPath, pngBytes);
        }

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Image insertion test (invalid URI handling):");

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting image tag. The -fitSize switch scales the image to the textbox size.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template to disk.
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportModel model = new();

        // -----------------------------------------------------------------
        // 3. Build the report with inline error messages enabled.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // BuildReport returns a bool indicating success when InlineErrorMessages is set.
        bool success = engine.BuildReport(reportDoc, model, "model");

        Console.WriteLine($"Report generation success flag: {success}");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
