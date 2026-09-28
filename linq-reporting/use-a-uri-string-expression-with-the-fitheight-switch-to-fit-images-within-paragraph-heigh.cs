using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // URI string that points to the image file.
    public string ImageUri { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Ensure the output folder exists.
        Directory.CreateDirectory("output");

        // Create a simple 1x1 pixel PNG image from a Base64 string.
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK0cAAAAASUVORK5CYII=";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            ImageUri = imagePath
        };

        // Create the LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a paragraph describing the image.
        builder.Writeln("Below is an image fitted to the paragraph height using the -fitHeight switch:");

        // Insert a textbox that will contain the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag with -fitHeight switch and URI expression.
        builder.Write("<<image [model.ImageUri] -fitHeight>>");

        // Save the template.
        string templatePath = Path.Combine("output", "template.docx");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        string outputPath = Path.Combine("output", "output.docx");
        reportDoc.Save(outputPath);
    }
}
