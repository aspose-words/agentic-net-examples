using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "ImageTemplate.docx";
        const string outputPath = "ImageReport.docx";

        // -------------------------------------------------
        // 1. Create the template document with LINQ tags.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting tag: image expression using a byte array and -fitSize switch.
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -------------------------------------------------
        var doc = new Document(templatePath);

        var model = new ReportModel(); // model.ImageData is pre‑initialized.

        // -------------------------------------------------
        // 3. Build the report using ReportingEngine.
        // -------------------------------------------------
        var engine = new ReportingEngine();
        // No special options are required for this scenario.
        engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        doc.Save(outputPath);
    }
}

// Data model used by the template.
// The ImageData property returns a byte array containing a tiny PNG image.
public class ReportModel
{
    // A 1x1 pixel transparent PNG encoded in Base64.
    private const string Base64Png =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";

    public byte[] ImageData { get; } = Convert.FromBase64String(Base64Png);
}
