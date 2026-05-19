using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // The image data provided as a stream.
    public Stream ImageStream { get; set; } = Stream.Null;
}

public class Program
{
    public static void Main()
    {
        // 1. Prepare a simple PNG image (1x1 pixel, red) as a base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9YVYB2cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // 2. Create the data model and assign a MemoryStream containing the image.
        var model = new ReportModel
        {
            ImageStream = new MemoryStream(imageBytes)
        };
        // Ensure the stream is positioned at the beginning before the engine reads it.
        model.ImageStream.Position = 0;

        // 3. Build the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a textbox that will host the image tag.
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting image tag with -fitSizeLim switch.
        builder.Write("<<image [model.ImageStream] -fitSizeLim>>");

        // 4. Generate the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the resulting document.
        doc.Save("ReportWithImage.docx");
    }
}
