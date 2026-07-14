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
        // 1. Prepare a folder for all generated files.
        // -----------------------------------------------------------------
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 2. Create a small PNG image from a Base64 string.
        // -----------------------------------------------------------------
        // This is a 1x1 pixel red PNG.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "b9XcAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        string imagePath = Path.Combine(workDir, "sample.png");
        File.WriteAllBytes(imagePath, imageBytes);

        // -----------------------------------------------------------------
        // 3. Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);

        // The -fitHeight switch scales the image to the height of the container
        // while preserving the original width proportionally.
        builder.Write("<<image [model.ImagePath] -fitHeight>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Load the template and run the reporting engine.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        ReportModel model = new ReportModel
        {
            ImagePath = imagePath // absolute path to the sample image.
        };

        ReportingEngine engine = new ReportingEngine();
        // The root object name must match the tag prefix used in the template.
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(workDir, "result.docx");
        loadedTemplate.Save(resultPath);

        // The example finishes without waiting for user input.
    }
}
