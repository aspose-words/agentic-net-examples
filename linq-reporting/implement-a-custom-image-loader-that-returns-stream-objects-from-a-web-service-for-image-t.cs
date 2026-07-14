using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document that contains an image tag inside a textbox.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var template = new Document(templatePath);

        // Prepare the data model.
        var model = new ReportModel
        {
            // The ImageUrl property is kept for compatibility but will not be used for downloading.
            ImageUrl = "https://via.placeholder.com/150"
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        template.Save("Report.docx");
    }

    // Creates a simple Word document with a textbox that hosts an image tag.
    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a textbox to host the image tag.
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // The image tag expects a Stream; the model provides it via ImageStream property.
        builder.Write("<<image [model.ImageStream] -fitSize>>");

        doc.Save(filePath);
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // URL of the image (kept for reference; not used for downloading in this example).
    public string ImageUrl { get; set; } = string.Empty;

    // Returns a fresh MemoryStream containing a small placeholder PNG image.
    // The stream is positioned at the beginning, as required by the engine.
    public Stream ImageStream
    {
        get
        {
            // Retrieve the image bytes from the helper (no network calls).
            var bytes = ImageService.DownloadImageBytes(ImageUrl).Result;
            var stream = new MemoryStream(bytes);
            stream.Position = 0;
            return stream;
        }
    }
}

// Helper that provides image bytes. In this example we embed a tiny PNG to avoid external calls.
public static class ImageService
{
    // Base64‑encoded 1×1 pixel PNG (transparent).
    private const string Base64Png =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9YVhZVIAAAAASUVORK5CYII=";

    // Returns the image bytes synchronously wrapped in a Task.
    public static Task<byte[]> DownloadImageBytes(string url)
    {
        // Decode the embedded PNG.
        byte[] imageBytes = Convert.FromBase64String(Base64Png);
        return Task.FromResult(imageBytes);
    }
}
