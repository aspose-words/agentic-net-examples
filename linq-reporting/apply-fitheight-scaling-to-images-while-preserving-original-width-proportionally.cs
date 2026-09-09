using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outDir);

        // -----------------------------------------------------------------
        // 1. Create a sample image file (a tiny red PNG).
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(outDir, "sample.png");
        // PNG data for a 2x2 red image.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFklEQVQImWNgYGD4z0AEYBxVSFIAAQAB" +
            "JwABX6Z1WQAAAABJRU5ErkJggg==");
        File.WriteAllBytes(imagePath, pngBytes);

        // -----------------------------------------------------------------
        // 2. Create the LINQ Reporting template.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag with -fitHeight switch. Width will be kept proportional.
        builder.Write("<<image [model.ImagePath] -fitHeight>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Data model exposing the image path.
        ReportModel model = new ReportModel
        {
            ImagePath = imagePath
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outDir, "Report.docx");
        reportDoc.Save(resultPath);
    }
}

// Data model used by the template.
public class ReportModel
{
    // Path to the image file that will be inserted.
    public string ImagePath { get; set; } = string.Empty;
}
