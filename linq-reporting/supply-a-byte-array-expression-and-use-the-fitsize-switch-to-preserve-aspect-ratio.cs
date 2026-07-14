using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a simple 1x1 PNG image as a byte array.
        var model = new ReportModel();

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert a textbox that will act as the image container.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Insert the image tag that uses the byte array expression and the -fitSize switch.
        builder.Writeln("<<image [model.ImageBytes] -fitSize>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        // Build the report using the model object; the root name in the template is "model".
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Public data model required by the template.
public class ReportModel
{
    // Byte array that holds the image data.
    public byte[] ImageBytes { get; set; }

    public ReportModel()
    {
        // A 1x1 pixel transparent PNG (base64 encoded).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        ImageBytes = Convert.FromBase64String(base64Png);
    }
}
