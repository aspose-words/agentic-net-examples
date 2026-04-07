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
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a simple image file that will be used in the report.
        // -----------------------------------------------------------------
        // This is a 1x1 pixel PNG (transparent) encoded in Base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9W8VYVQAAAAASUVORK5CYII=";
        string imagePath = Path.Combine(outputDir, "sample.png");
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Move the cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the LINQ Reporting image tag with the -fitWidth switch.
        // The expression [model.ImagePath] will be resolved by the ReportingEngine.
        builder.Write("<<image [model.ImagePath] -fitWidth>>");

        // Return to the main document body.
        builder.MoveToDocumentEnd();

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            ImagePath = imagePath // Full path to the image file.
        };

        // -----------------------------------------------------------------
        // 4. Load the template and build the report.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        // The root object name in the template is "model".
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "ReportWithHeaderImage.docx");
        loadedTemplate.Save(resultPath);

        // Inform the user (no interactive input required).
        Console.WriteLine("Report generated at: " + resultPath);
    }
}

// Simple data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Path to the image that will be inserted into the header.
    public string ImagePath { get; set; } = string.Empty;
}
