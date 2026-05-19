using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample PNG image file.
        string imagePath = Path.Combine(outputDir, "sample.png");
        CreateSamplePng(imagePath);

        // Create the LINQ Reporting template with an image tag inside a textbox.
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath, imagePath);

        // Load the template document.
        Document templateDoc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            ImagePath = imagePath
        };

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        templateDoc.Save(resultPath);
    }

    // Generates a minimal 1x1 pixel PNG image.
    private static void CreateSamplePng(string filePath)
    {
        // PNG data for a 1x1 red pixel.
        byte[] pngData = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "Z6VYAAAAASUVORK5CYII=");
        File.WriteAllBytes(filePath, pngData);
    }

    // Creates a DOCX template containing a textbox with an image tag.
    private static void CreateTemplate(string filePath, string imagePlaceholder)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox that will host the image.
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // The image tag uses the model's ImagePath property.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template.
        doc.Save(filePath);
    }

    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        public string ImagePath { get; set; } = string.Empty;
    }
}
