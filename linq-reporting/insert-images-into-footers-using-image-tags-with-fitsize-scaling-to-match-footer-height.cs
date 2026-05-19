using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template, image and final report.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string imagePath = Path.Combine(workDir, "sample.png");
        string reportPath = Path.Combine(workDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple PNG image (1x1 pixel, red) from a Base64 string.
        // -----------------------------------------------------------------
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YVYcVYAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, pngBytes);

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Move to the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 30);
        builder.MoveTo(textBox.FirstParagraph);

        // Write the LINQ Reporting image tag with -fitSize switch.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and populate it using ReportingEngine.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Data model exposing the image path.
        var model = new ReportModel { ImagePath = imagePath };

        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class ReportModel
{
    public string ImagePath { get; set; } = string.Empty;
}
