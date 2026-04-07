using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a simple data model with an image stored as a byte array.
        ReportModel model = new ReportModel
        {
            // A 1x1 pixel PNG image (transparent) encoded in Base64.
            ImageBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=")
        };

        // Build the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the LINQ Reporting image tag using the byte array expression and -fitSize switch.
        builder.Write("<<image [model.ImageBytes] -fitSize>>");

        // Save the template (optional, shown for completeness).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The root object name in the template is "model".
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        template.Save(outputPath);
    }
}

// Public data model required by the template.
public class ReportModel
{
    // Byte array containing image data.
    public byte[] ImageBytes { get; set; } = Array.Empty<byte>();
}
