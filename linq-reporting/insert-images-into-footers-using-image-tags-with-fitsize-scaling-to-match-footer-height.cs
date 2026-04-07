using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Folder for generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple PNG image (1x1 red pixel) and save it locally.
        // The image data is a PNG byte array.
        byte[] pngData = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
            0xDE,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,
            0x00,0x04,0x00,0x01,0xE2,0x26,0x05,0x9B,
            0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
            0xAE,0x42,0x60,0x82
        };
        string imagePath = Path.Combine(outputDir, "footer.png");
        File.WriteAllBytes(imagePath, pngData);

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            FooterImagePath = imagePath
        };

        // Create the template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Move cursor to the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert a textbox that will contain the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 50);
        // Position the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the LINQ Reporting image tag with -fitSize switch.
        builder.Write("<<image [model.FooterImagePath] -fitSize>>");

        // Return to the main body (optional).
        builder.MoveToDocumentEnd();
        builder.Writeln("Report body content.");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report using LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        string resultPath = Path.Combine(outputDir, "ReportWithFooterImage.docx");
        reportDoc.Save(resultPath);
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Path to the image that will be placed in the footer.
    public string FooterImagePath { get; set; } = string.Empty;
}
