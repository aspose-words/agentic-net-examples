using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create a small red dot PNG image from a Base64 string.
        string imagePath = Path.Combine(outputDir, "sample.png");
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFklEQVQoU2NkYGD4z0AEYBxVSFIAAQAA//8DAF6cB6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, imageBytes);

        // Build the data model. Use a full absolute path so the engine can locate the image.
        ReportModel model = new()
        {
            ImagePath = Path.GetFullPath(imagePath)
        };

        // Create the template document with a header that contains an image tag inside a textbox.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Move to the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a textbox shape that will hold the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 100);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag with -fitWidth scaling switch.
        builder.Write("<<image [model.ImagePath] -fitWidth>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new(templatePath);

        // Build the report using the LINQ Reporting Engine.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        string outputPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(outputPath);
    }
}

// Public data model class.
public class ReportModel
{
    // Absolute path to the image file to be inserted.
    public string ImagePath { get; set; } = string.Empty;
}
