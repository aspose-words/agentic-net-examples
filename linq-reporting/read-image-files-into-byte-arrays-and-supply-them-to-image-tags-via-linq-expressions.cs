using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a sample image file (a tiny red dot PNG) and read it.
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(workDir, "sample.png");
        // Base64 for a 1x1 red PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9WcVhVQAAAAASUVORK5CYII=";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // Load the image bytes into the data model.
        ReportModel model = new ReportModel
        {
            ImageData = File.ReadAllBytes(imagePath),
            Title = "Sample Image"
        };

        // ---------------------------------------------------------------
        // 2. Build the LINQ Reporting template programmatically.
        // ---------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a title.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag must be inside the textbox. Use -fitSize switch to fit the image.
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Load the template and build the report.
        // ---------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        // BuildReport with root name "model" to match the tags.
        engine.BuildReport(reportDoc, model, "model");

        // ---------------------------------------------------------------
        // 4. Save the final document.
        // ---------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "Report.docx");
        reportDoc.Save(outputPath);

        // The example finishes without waiting for user input.
    }
}

// Public data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Image data supplied as a byte array.
    public byte[] ImageData { get; set; } = Array.Empty<byte>();

    // Simple text to demonstrate a regular expression tag.
    public string Title { get; set; } = string.Empty;
}
