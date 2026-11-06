using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working directories
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string outputPath = Path.Combine(workDir, "Report.docx");
        string imagePath = Path.Combine(workDir, "sample.png");

        // Create a tiny PNG image (1x1 pixel, red) and save it locally
        // Base64 for a 1x1 red PNG
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/6VZcWQAAAABJRU5ErkJggg==";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, imageBytes);

        // -------------------------
        // Create the LINQ Reporting template
        // -------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a paragraph describing the image
        builder.Writeln("Below is an image stretched to fill the textbox using -fitSize:");

        // Insert a textbox shape that will contain the image tag
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        // Move cursor inside the textbox
        builder.MoveTo(textBox.FirstParagraph);
        // Write the image tag with -fitSize switch
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------
        // Load the template for report generation
        // -------------------------
        Document reportDoc = new Document(templatePath);

        // Prepare the data model
        ReportModel model = new ReportModel
        {
            ImagePath = imagePath
        };

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        reportDoc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting engine
public class ReportModel
{
    // Path to the image file that will be inserted into the document
    public string ImagePath { get; set; } = string.Empty;
}
