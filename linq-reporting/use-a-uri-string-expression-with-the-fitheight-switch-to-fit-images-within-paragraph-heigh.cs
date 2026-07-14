using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Path to the image that will be inserted into the report.
        public string ImagePath { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Ensure the output folder exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample image file that will be referenced from the template.
        // -----------------------------------------------------------------
        string imageFile = Path.Combine(outputDir, "sample.png");
        CreateSamplePng(imageFile);

        // -----------------------------------------------------------------
        // 2. Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and run the ReportingEngine.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Prepare the data source.
        ReportModel model = new ReportModel { ImagePath = imageFile };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "result.docx");
        doc.Save(resultPath);
    }

    // Creates a minimal PNG image (1x1 pixel, red) and saves it to the specified path.
    private static void CreateSamplePng(string filePath)
    {
        // PNG binary for a 1x1 red pixel.
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
        File.WriteAllBytes(filePath, pngData);
    }

    // Generates a Word document that contains a textbox with an image tag using the -fitHeight switch.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph that will contain the textbox.
        builder.Writeln("Image inside a textbox (fitHeight):");

        // Insert a textbox shape that will act as the image container.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting image tag with -fitHeight switch.
        builder.Write("<<image [model.ImagePath] -fitHeight>>");

        // Save the template.
        doc.Save(filePath);
    }
}
