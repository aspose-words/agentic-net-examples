using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class ReportModel
{
    public string Title { get; set; } = "Sample Image Report";
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string imagePath = Path.Combine(workDir, "sample.png");
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "output.docx");

        // Create a minimal 1x1 PNG image from a Base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, pngBytes);

        // Load the image bytes for the data model.
        byte[] imageBytes = File.ReadAllBytes(imagePath);
        var model = new ReportModel { ImageData = imageBytes };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title placeholder.
        builder.Writeln("<<[model.Title]>>");

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Save the template, then reload it as required by the workflow.
        templateDoc.Save(templatePath);
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the final document.
        loadedTemplate.Save(outputPath);
    }
}
