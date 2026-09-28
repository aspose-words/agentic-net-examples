using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template document.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will hold the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag expects a byte array; the model will provide it.
        builder.Write("<<image [model.ImageBytes] -fitSize>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "template.docx");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample data with a Base64‑encoded image.
        // -----------------------------------------------------------------
        // This is a 1x1 red PNG pixel.
        const string base64RedPixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BFwAE/wJ/lGkAAAAASUVORK5CYII=";
        ReportModel model = new ReportModel
        {
            Base64Image = base64RedPixel
        };

        // -----------------------------------------------------------------
        // 4. Build the report.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // No special options required for this example.
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(outputDir, "result.docx");
        doc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Base64 string representing the image.
    public string Base64Image { get; set; } = string.Empty;

    // Byte array derived from the Base64 string; used by the <<image>> tag.
    public byte[] ImageBytes => Convert.FromBase64String(Base64Image);
}
