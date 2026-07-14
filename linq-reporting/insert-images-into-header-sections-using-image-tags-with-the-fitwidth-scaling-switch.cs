using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Path to the image file that will be inserted into the header.
    public string ImagePath { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Prepare an output folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a simple PNG image (1x1 transparent pixel).
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(workDir, "SampleImage.png");
        CreateSamplePng(imagePath);

        // -----------------------------------------------------------------
        // 2. Build the template document that contains an image tag in the header.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "Template.docx");
        CreateTemplate(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var model = new ReportModel { ImagePath = imagePath };

        // -----------------------------------------------------------------
        // 4. Run the LINQ Reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string reportPath = Path.Combine(workDir, "Report.docx");
        doc.Save(reportPath);
    }

    // Creates a minimal 1x1 PNG image from a Base64 string.
    private static void CreateSamplePng(string filePath)
    {
        // Base64 for a 1x1 transparent PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lKX9WQAAAABJRU5ErkJggg==";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(filePath, pngBytes);
    }

    // Creates a Word template with a header that contains an image tag using the -fitWidth switch.
    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Move to the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting image tag with -fitWidth scaling switch.
        builder.Write("<<image [model.ImagePath] -fitWidth>>");

        // Return to the main document body.
        builder.MoveToDocumentEnd();

        // Save the template.
        doc.Save(filePath);
    }
}
