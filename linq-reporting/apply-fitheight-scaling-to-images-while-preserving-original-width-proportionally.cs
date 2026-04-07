using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create a simple PNG image (1x1 pixel, red) and save it locally.
        // The image data is stored as a Base64 string to avoid external dependencies.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        string imagePath = Path.Combine(workDir, "sample.png");
        File.WriteAllBytes(imagePath, imageBytes);

        // 2. Build the LINQ Reporting template programmatically.
        // The image tag uses the -fitHeight switch to scale the image height to the textbox height,
        // while preserving the original width proportionally.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will act as the image container.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox so we can write the image tag.
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting image tag with fitHeight scaling.
        builder.Write("<<image [model.ImagePath] -fitHeight>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // 4. Prepare the data model. The property name must match the tag expression.
        ReportModel model = new ReportModel
        {
            ImagePath = imagePath
        };

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this scenario.
        engine.BuildReport(reportDoc, model, "model");

        // 6. Save the generated document.
        string outputPath = Path.Combine(workDir, "output.docx");
        reportDoc.Save(outputPath);
    }
}

// Public data model class required by the LINQ Reporting engine.
public class ReportModel
{
    // Path to the image file that will be inserted into the document.
    public string ImagePath { get; set; } = string.Empty;
}
