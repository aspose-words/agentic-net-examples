using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Stream containing image data. Initialized in the constructor.
    public Stream ImageStream { get; set; } = Stream.Null;
}

public class Program
{
    public static void Main()
    {
        // Sample PNG image (1x1 pixel, transparent) encoded in base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAn8B9pVYVQAAAABJRU5ErkJggg==";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        var imageStream = new MemoryStream(pngBytes);
        // Ensure the stream is positioned at the beginning.
        imageStream.Position = 0;

        // Prepare the data model.
        var model = new ReportModel
        {
            ImageStream = imageStream
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a textbox to host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag using a Stream source and the -fitSizeLim switch.
        builder.Write("<<image [model.ImageStream] -fitSizeLim>>");

        // Reset the stream before the reporting engine processes it.
        model.ImageStream.Position = 0;

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
