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
        Directory.CreateDirectory("output");

        // Create a simple PNG image (1x1 red pixel) and save it locally.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        File.WriteAllBytes(imagePath, imageBytes);

        // Create the template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Move to the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert a textbox shape that will contain the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 30);
        // Position the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the image tag with -fitSize switch.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template to disk.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new()
        {
            ImagePath = imagePath
        };

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output", "ReportWithFooterImage.docx");
        doc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting template.
public class ReportModel
{
    // Path to the image that will be inserted into the footer.
    public string ImagePath { get; set; } = string.Empty;
}
