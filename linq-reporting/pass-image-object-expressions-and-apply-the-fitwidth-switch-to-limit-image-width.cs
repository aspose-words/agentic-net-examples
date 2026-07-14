using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare working folder and file paths.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string imagePath = Path.Combine(workDir, "sample.png");
        string outputPath = Path.Combine(workDir, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple PNG image file (1x1 pixel) that will be used in the report.
        // -----------------------------------------------------------------
        // Base64-encoded PNG (transparent 1x1 pixel).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, pngBytes);

        // -----------------------------------------------------------------
        // 2. Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag uses a byte[] expression and applies the -fitWidth switch.
        builder.Write("<<image [model.ImageData] -fitWidth>>");

        // Save the template to disk (required before BuildReport).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model containing the image bytes.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            ImageData = pngBytes
        };

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);

        // Clean up temporary files.
        File.Delete(templatePath);
        File.Delete(imagePath);
    }
}

// ---------------------------------------------------------------------
// Data model used by the LINQ Reporting template.
// The property must be public and non‑nullable to avoid warnings.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Image data as a byte array; supported by the image tag.
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}
