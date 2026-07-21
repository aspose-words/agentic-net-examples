using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the image tag that uses a byte[] expression and the -fitSize switch.
        builder.Write("<<image [model.ImageBytes] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Sample PNG image (1x1 pixel) encoded as Base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lZ0ZAAAAAElFTkSuQmCC";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        var model = new ReportModel
        {
            ImageBytes = imageBytes
        };

        // -----------------------------------------------------------------
        // 3. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        // The root object name in the template is "model".
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(reportPath);
    }
}

// Data model used by the template. The property is initialized to avoid nullable warnings.
public class ReportModel
{
    public byte[] ImageBytes { get; set; } = Array.Empty<byte>();
}
