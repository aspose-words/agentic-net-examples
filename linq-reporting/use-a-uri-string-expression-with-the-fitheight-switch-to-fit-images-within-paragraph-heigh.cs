using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a simple PNG image (1x1 pixel, red) and save it locally.
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        if (!File.Exists(imagePath))
        {
            // Base64 for a 1x1 red PNG.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGNgYGBgAAAABQABDQottAAAAABJRU5ErkJggg==";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imagePath, imageBytes);
        }

        // Create the data model that will be bound to the template.
        var model = new ReportModel
        {
            ImageUri = imagePath // URI (file path) used by the image tag.
        };

        // -----------------------------------------------------------------
        // Step 1: Build the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Insert the LINQ Reporting image tag with the -fitHeight switch.
        builder.Write("<<image [model.ImageUri] -fitHeight>>");

        // Add a regular paragraph after the textbox for visual reference.
        builder.MoveToDocumentEnd();
        builder.Writeln("Image fitted to the height of the containing paragraph.");

        // Save the template to disk.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // Default options.
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // Step 3: Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "report.docx");
        doc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Path or URI to the image that will be inserted.
    public string ImageUri { get; set; } = string.Empty;
}
