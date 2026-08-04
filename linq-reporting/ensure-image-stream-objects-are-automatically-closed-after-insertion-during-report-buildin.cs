using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template with a textbox that contains an image tag.
        var templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Prepare the data model that supplies an image stream.
        var model = new ReportModel();

        // Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options required.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var reportPath = "Report.docx";
        doc.Save(reportPath);
        Console.WriteLine($"Report saved to '{reportPath}'.");

        // Verify that the image stream has been closed automatically.
        try
        {
            // Accessing a closed stream throws ObjectDisposedException.
            var _ = model.ImageStream.Position;
            Console.WriteLine("Stream is still open (unexpected).");
        }
        catch (ObjectDisposedException)
        {
            Console.WriteLine("Image stream has been automatically closed after report building.");
        }
    }

    // Creates a Word document containing a textbox with an image tag.
    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting tag that inserts the image from the stream.
        builder.Write("<<image [model.ImageStream] -fitSize>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // The image stream is initialized with a tiny PNG image.
    public Stream ImageStream { get; set; }

    public ReportModel()
    {
        // A 1x1 pixel transparent PNG (base64 encoded).
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9YVh" +
            "XwAAAABJRU5ErkJggg==";

        byte[] pngBytes = Convert.FromBase64String(base64Png);
        ImageStream = new MemoryStream(pngBytes);
    }
}
