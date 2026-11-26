using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // URL of the image to fetch from a web service.
    public string ImageUrl { get; set; } = "https://via.placeholder.com/150";

    // Cached image bytes to avoid multiple HTTP calls.
    private byte[]? _imageBytes;

    // Returns a fresh Stream each time; LINQ Reporting will consume it.
    public Stream ImageStream
    {
        get
        {
            if (_imageBytes == null)
            {
                using var http = new HttpClient();
                // Synchronously fetch the image bytes.
                _imageBytes = http.GetByteArrayAsync(ImageUrl).Result;
            }
            // Return a new MemoryStream positioned at the start.
            return new MemoryStream(_imageBytes);
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "template.docx";
        const string outputPath = "report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report with image loaded from a web service:");

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting image tag referencing the ImageStream property.
        builder.Write("<<image [model.ImageStream] -fitSize>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var template = new Document(templatePath);
        var model = new ReportModel();

        var engine = new ReportingEngine();
        // BuildReport processes the image tag and fetches the stream.
        engine.BuildReport(template, model, "model");

        // Save the final report.
        template.Save(outputPath);
    }
}
