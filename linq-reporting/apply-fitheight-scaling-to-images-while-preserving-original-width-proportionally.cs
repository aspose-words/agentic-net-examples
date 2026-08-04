using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create a simple image file (1x1 pixel PNG) from a Base64 string.
        string imagePath = Path.Combine(workDir, "sample.png");
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YVh3V8AAAAASUVORK5CYII=";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // 2. Build the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a title.
        builder.Writeln("Image scaling with -fitHeight:");

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // The image tag uses -fitHeight to fit the image height while preserving width proportionally.
        builder.Write("<<image [model.ImagePath] -fitHeight>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Prepare the data model.
        ReportModel model = new ReportModel
        {
            ImagePath = imagePath
        };

        // 4. Load the template and build the report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the generated report.
        string reportPath = Path.Combine(workDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}

// Simple data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Path to the image that will be inserted into the document.
    public string ImagePath { get; set; } = string.Empty;
}
