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
        const string outputFolder = "Output";
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // 1. Create a sample image (1x1 red pixel) and save it to disk.
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(outputFolder, "sample.png");
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/6V4bAAAAAElFTkSuQmCC";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // -----------------------------------------------------------------
        // 2. Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputFolder, "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // The -fitHeight switch scales the image to the textbox height while keeping the original width proportionally.
        builder.Write("<<image [model.ImagePath] -fitHeight>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Data model referenced by the template.
        var model = new ReportModel
        {
            // Use an absolute path so the engine can locate the image reliably.
            ImagePath = Path.GetFullPath(imagePath)
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputFolder, "Report.docx");
        reportDoc.Save(resultPath);
    }
}

// ---------------------------------------------------------------------
// Data model used by the LINQ Reporting template.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Path to the image file that will be inserted into the report.
    public string ImagePath { get; set; } = string.Empty;
}
