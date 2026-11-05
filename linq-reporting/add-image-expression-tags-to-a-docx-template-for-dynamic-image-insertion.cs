using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using System.Text;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample PNG image (1x1 pixel, red).
        string imagePath = Path.Combine(outputDir, "sample.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "Zc6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // -----------------------------------------------------------------
        // Step 1: Create the LINQ Reporting template with an image tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will hold the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "template.docx");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template for report generation.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Step 3: Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            ImagePath = imagePath
        };

        // -----------------------------------------------------------------
        // Step 4: Build the report using LINQ Reporting Engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // Step 5: Save the generated report.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "report.docx");
        doc.Save(resultPath);
    }
}

// Data model used by the template.
public class ReportModel
{
    // Path to the image file that will be inserted.
    public string ImagePath { get; set; } = string.Empty;
}
