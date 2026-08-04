using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Path to the image that will be inserted by the LINQ Reporting engine.
    public string ImagePath { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Working directory for all temporary files.
        string workDir = Directory.GetCurrentDirectory();

        // -----------------------------------------------------------------
        // 1. Create a sample image file (a tiny red PNG) that will be used
        //    by the report. The image data is stored as a Base64 string.
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(workDir, "sample.png");
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BFwAE/wJ/6V4AAAAASUVORK5CYII=";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // -----------------------------------------------------------------
        // 2. Build the LINQ Reporting template programmatically.
        //    The image tag must be placed inside a TextBox shape and use
        //    the -fitSize switch to stretch the image to fill the paragraph.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a TextBox that will act as the image container.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        // Move the cursor inside the TextBox.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the image tag. The expression refers to the model's ImagePath property.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportModel model = new ReportModel { ImagePath = imagePath };

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "output.docx");
        doc.Save(outputPath);

        // Inform the user (optional, no interactive input required).
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
