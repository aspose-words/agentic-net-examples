using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Path to the image that will be inserted dynamically.
        public string ImagePath { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare working folders.
        // -----------------------------------------------------------------
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 2. Create a tiny PNG image (1x1 pixel, red) and save it locally.
        // -----------------------------------------------------------------
        // Base64 representation of a 1x1 red PNG.
        const string pngBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lK5XAAAAAElFTkSuQmCC";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);
        string imagePath = Path.Combine(workDir, "sample.png");
        File.WriteAllBytes(imagePath, pngBytes);

        // -----------------------------------------------------------------
        // 3. Build the DOCX template programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple heading.
        builder.Writeln("Report with dynamic image insertion:");

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the LINQ Reporting image tag. The expression refers to the model's ImagePath property.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Load the template and populate it using the ReportingEngine.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            ImagePath = imagePath // The tag will resolve this path at runtime.
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Optional: you can check the success flag if you enabled InlineErrorMessages.
        if (!success)
        {
            Console.WriteLine("Report generation encountered errors.");
        }

        // -----------------------------------------------------------------
        // 5. Save the final document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "Report.docx");
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated successfully: {outputPath}");
    }
}
