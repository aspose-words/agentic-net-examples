using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Sample Base64 string representing a 1x1 pixel PNG image.
        const string sampleBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ+XKcAAAAASUVORK5CYII=";

        // Prepare the data model.
        var model = new ReportModel
        {
            ImageBase64 = sampleBase64
        };

        // Create a new blank document and a builder to add content.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting image tag – the expression returns a byte[].
        builder.Write("<<image [model.ImageBytes] -fitSize>>");

        // Build the report using the model as the root data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("ReportWithImage.docx");
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Base64 representation of the image.
    public string ImageBase64 { get; set; } = string.Empty;

    // Byte array obtained from the Base64 string – used by the image tag.
    public byte[] ImageBytes => Convert.FromBase64String(ImageBase64);
}
