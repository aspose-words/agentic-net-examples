using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Image data as a byte array – supported by the LINQ Reporting image tag.
    public byte[] Photo { get; set; }

    public ReportModel(byte[] photo)
    {
        Photo = photo;
    }
}

public class Program
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "report.docx");
        string imagePath = Path.Combine(workDir, "sample.png");

        // Create a minimal PNG image (1x1 pixel) and save it.
        // Base64 for a transparent 1x1 PNG.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BFwAE/wJ/6V9sAAAAAElFTkSuQmCC";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, pngBytes);

        // Load the image bytes for the model.
        byte[] imgData = File.ReadAllBytes(imagePath);
        ReportModel model = new ReportModel(imgData);

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will contain the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag with -fitWidth switch to limit image width to the textbox width.
        builder.Write("<<image [model.Photo] -fitWidth>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save(outputPath);
    }
}
