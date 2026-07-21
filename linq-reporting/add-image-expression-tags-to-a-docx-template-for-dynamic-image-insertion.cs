using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for files used in the example.
        const string templatePath = "template.docx";
        const string imagePath = "sample.png";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple PNG image (1x1 pixel, red) and save it locally.
        // -----------------------------------------------------------------
        // The image data is a base64‑encoded PNG.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "ZcKcAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, pngBytes);

        // -----------------------------------------------------------------
        // 2. Build the DOCX template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report with dynamic image:");
        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Insert the LINQ Reporting image tag.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model that supplies the image path.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            ImagePath = Path.GetFullPath(imagePath) // absolute path works reliably.
        };

        // -----------------------------------------------------------------
        // 4. Load the template and run the reporting engine.
        // -----------------------------------------------------------------
        Document docToReport = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        // Build the report; the root object name must match the tag prefix ("model").
        engine.BuildReport(docToReport, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        docToReport.Save(outputPath);
    }
}

// Simple data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Path (or URI) to the image that will be inserted.
    public string ImagePath { get; set; } = string.Empty;
}
