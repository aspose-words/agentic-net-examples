using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Path to the image that will be inserted into the footer.
        public string ImagePath { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Ensure the output folder exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a tiny PNG image (1x1 pixel, red) from a Base64 string.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9WcK4VQAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        string imagePath = Path.Combine(outputDir, "sample.png");
        File.WriteAllBytes(imagePath, imageBytes);

        // Build the LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath, imagePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel { ImagePath = imagePath };

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        string resultPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(resultPath);
    }

    // Creates a Word document that contains a footer with an image tag.
    private static void CreateTemplate(string filePath, string placeholderImagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert a textbox that will hold the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 30);
        // Ensure the cursor is inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);

        // Write the LINQ Reporting image tag with -fitSize switch.
        // The tag will be replaced with the actual image during report generation.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Optionally add some static text after the image.
        builder.Writeln(" Footer text");

        // Save the template.
        doc.Save(filePath);
    }
}
