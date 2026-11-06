using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a simple PNG image as a byte array (a 1x1 red pixel).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK7cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Create the data model with a public byte[] property.
        var model = new ReportModel
        {
            ImageBytes = imageBytes
        };

        // Build the template document.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert a textbox that will contain the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting image tag with -fitSize switch to preserve aspect ratio.
        builder.Write("<<image [model.ImageBytes] -fitSize>>");

        // Save the template (optional, demonstrates lifecycle).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back (simulating a real scenario).
        var loadedTemplate = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the final report.
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);
    }

    // Public data model class required by the reporting engine.
    public class ReportModel
    {
        // Image data supplied as a byte array.
        public byte[] ImageBytes { get; set; } = Array.Empty<byte>();
    }
}
