using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

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
        // -----------------------------------------------------------------
        // Prepare working folders.
        // -----------------------------------------------------------------
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // Create a sample image file (a tiny red pixel) that will be used
        // by the report. The image is stored as a PNG byte array.
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(workDir, "sample.png");
        // Valid Base64 for a 1x1 red PNG.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BFwAI/wN/6XcKAAAAAElFTkSuQmCC");
        File.WriteAllBytes(imagePath, pngBytes);

        // -----------------------------------------------------------------
        // Build the LINQ Reporting template programmatically.
        // The template contains a textbox with an image tag that uses the -fitSize switch.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will act as the image container.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        // Move the cursor inside the textbox's first paragraph.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the image tag. The expression [model.ImagePath] will be resolved from the data source.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Prepare the data source.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel { ImagePath = imagePath };

        // -----------------------------------------------------------------
        // Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "Report.docx");
        reportDoc.Save(outputPath);

        // Indicate successful completion.
        Console.WriteLine("Report generated at: " + outputPath);
    }
}
