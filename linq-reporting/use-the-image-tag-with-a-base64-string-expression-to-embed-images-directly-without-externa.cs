using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox – the required container for an image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the LINQ Reporting image tag that will receive a Base64 string.
        builder.Write("<<image [model.ImageBase64] -fitSize>>");

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            // A 1x1 pixel transparent PNG encoded as Base64.
            ImageBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X3V8AAAAASUVORK5CYII="
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportWithImage.docx");
        template.Save(outputPath);
    }
}

// Simple public data model required by the template.
public class ReportModel
{
    // The Base64 representation of the image to embed.
    public string ImageBase64 { get; set; } = string.Empty;
}
