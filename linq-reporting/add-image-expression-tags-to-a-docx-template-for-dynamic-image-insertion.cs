using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample PNG image (1x1 pixel, red) from a Base64 string.
        string imagePath = Path.Combine(outputDir, "sample.png");
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YV6vV8AAAAASUVORK5CYII=";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // 2. Build the template document programmatically.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report with dynamic image:");
        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the LINQ Reporting image tag. The expression refers to the model's ImagePath property.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Prepare the data model.
        ReportModel model = new ReportModel
        {
            ImagePath = imagePath // Path to the image file created above.
        };

        // 5. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 6. Save the final document.
        string resultPath = Path.Combine(outputDir, "result.docx");
        reportDoc.Save(resultPath);
    }
}

// Public data model used by the template.
public class ReportModel
{
    // Path to the image that will be inserted.
    public string ImagePath { get; set; } = string.Empty;
}
