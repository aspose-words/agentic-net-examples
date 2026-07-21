using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Stream source for the image. Initialized to avoid nullable warnings.
    public Stream Image { get; set; }

    public ReportModel(Stream image)
    {
        Image = image;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a simple 1x1 PNG image from a Base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        using var imageStream = new MemoryStream(imageBytes);
        imageStream.Position = 0; // Ensure the stream is at the beginning.

        // Prepare the data model with the image stream.
        var model = new ReportModel(imageStream);

        // Build the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a textbox that will contain the image tag.
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Use the fitSizeLim switch to limit image scaling while preserving aspect ratio.
        builder.Write("<<image [model.Image] -fitSizeLim>>");

        // Reset the stream before the reporting engine consumes it.
        model.Image.Position = 0;

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("ReportWithFitSizeLim.docx");
    }
}
