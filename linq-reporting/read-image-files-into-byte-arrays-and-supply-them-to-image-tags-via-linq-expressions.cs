using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "work");
        Directory.CreateDirectory(workDir);

        // Create a minimal PNG (1x1 pixel, transparent) and write it to a file.
        // This binary data represents a valid PNG image.
        byte[] pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
            0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
            0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82
        };
        string imagePath = Path.Combine(workDir, "sample.png");
        File.WriteAllBytes(imagePath, pngBytes);

        // Load the image file into a byte array for the data model.
        byte[] imageData = File.ReadAllBytes(imagePath);

        // Data model used by the LINQ Reporting engine.
        ReportModel model = new ReportModel
        {
            Title = "Sample Image Report",
            ImageData = imageData
        };

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title paragraph.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln(); // empty line.

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag – the expression returns a byte[].
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Save the template to disk.
        string templatePath = Path.Combine(workDir, "template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the final document.
        string outputPath = Path.Combine(workDir, "ReportWithImage.docx");
        report.Save(outputPath);
    }
}

// Public data model with non‑nullable properties initialized.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}
