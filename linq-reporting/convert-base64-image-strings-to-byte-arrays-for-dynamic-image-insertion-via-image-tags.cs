using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Base64 string of a small PNG image (1x1 pixel)
    public string Base64Image { get; set; } = string.Empty;

    // Convert Base64 string to byte array for the image tag
    public byte[] ImageData => Convert.FromBase64String(Base64Image);
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data with a Base64 encoded image
        var model = new ReportModel
        {
            Base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XGkUAAAAASUVORK5CYII="
        };

        // Create a template document programmatically
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a textbox to host the image tag
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Save the template to disk
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting
        var loadedTemplate = new Document(templatePath);

        // Build the report using LINQ Reporting Engine
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);
    }
}
