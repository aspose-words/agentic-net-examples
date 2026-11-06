using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Image data that will be inserted into the report.
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}

public class Program
{
    public static void Main()
    {
        // File names used in the example.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";
        const string imagePath = "sample.png";

        // Create a tiny PNG image (1x1 pixel) from a Base64 string.
        // This avoids the need for System.Drawing dependencies.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Build the LINQ Reporting template.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag with -fitWidth switch to limit the image width.
        builder.Write("<<image [model.ImageData] -fitWidth>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Prepare the model with the image data.
        ReportModel model = new ReportModel { ImageData = File.ReadAllBytes(imagePath) };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(outputPath);

        // Optionally, you could check the success flag here.
        // The program ends here.
    }
}
