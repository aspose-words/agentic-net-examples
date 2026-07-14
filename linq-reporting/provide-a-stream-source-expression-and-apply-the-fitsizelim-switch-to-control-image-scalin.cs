using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Stream that provides the image data for the report.
    public Stream ImageStream { get; set; } = null!;
}

public class Program
{
    public static void Main()
    {
        // Create a simple PNG image (1x1 pixel, red) as a base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Prepare the data model with a reset MemoryStream.
        var model = new ReportModel
        {
            ImageStream = new MemoryStream(imageBytes, writable: false)
        };
        model.ImageStream.Position = 0; // Ensure the stream is at the beginning.

        // Build the template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a textbox that will host the image.
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting tag: image from a stream with -fitSizeLim switch.
        builder.Write("<<image [model.ImageStream] -fitSizeLim>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("ReportWithFitSizeLim.docx");
    }
}
