using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample PNG file (a tiny 1×1 red pixel) from a Base64 string
        string imagePath = Path.Combine(outputDir, "sample.png");
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, pngBytes);

        // Build the template document containing a textbox with an image tag
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox (shape) to host the image tag
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 300);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Writeln("<<image [model.Photo]>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Prepare the data model – expose the image as a byte array
        ReportModel model = new ReportModel(imagePath);

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report
        string resultPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(resultPath);
    }
}

// Data model exposing the image as a byte array (supported by the image tag)
public class ReportModel
{
    public byte[] Photo { get; }

    public ReportModel(string imageUri)
    {
        // Load the image file into a byte array
        Photo = File.ReadAllBytes(imageUri);
    }
}
