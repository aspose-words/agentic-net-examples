using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a simple image file that will be used in the report.
        //    The image is a 1x1 pixel PNG (transparent) encoded in Base64.
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(workDir, "sample.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        //    The image tag is placed inside a textbox so that the
        //    -fitHeight switch can control its size relative to the paragraph.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a paragraph that will contain the textbox.
        builder.Writeln("Below is an image fitted to the paragraph height:");

        // Insert a textbox shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting tag: image expression with -fitHeight switch.
        builder.Write("<<image [model.ImageUri] -fitHeight>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Data model for the report.
        ReportModel model = new ReportModel
        {
            ImageUri = imagePath // Path to the image file.
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the final document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "Report.docx");
        reportDoc.Save(outputPath);
    }
}

// Simple data model used by the template.
public class ReportModel
{
    // URI or file path to the image that will be inserted.
    public string ImageUri { get; set; } = string.Empty;
}
