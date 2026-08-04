using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ImageModel
{
    // Stream containing the image data. Initialized with a 1x1 PNG.
    public Stream ImageStream { get; set; } = new MemoryStream(Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
        "Z6VQAAAAASUVORK5CYII="));
}

public class Program
{
    public static void Main()
    {
        // Ensure the image stream is positioned at the beginning before the engine reads it.
        var model = new ImageModel();
        model.ImageStream.Position = 0;

        // Create a blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox to host the image tag (required by LINQ Reporting).
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting tag: insert the image from the stream and limit its size.
        builder.Write("<<image [model.ImageStream] -fitSizeLim>>");

        // Build the report using the model as the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated document.
        template.Save("ReportWithImage.docx");
    }
}
