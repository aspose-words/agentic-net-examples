using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Holds the Base64 image data (without the data URI prefix) that will be inserted via the LINQ Reporting engine.
    public string ImageBase64 { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // File names for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will contain the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting image tag – the expression returns a Base64 string (no data URI prefix).
        builder.Write("<<image [model.ImageBase64] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Prepare a Base64‑encoded PNG image (without the data URI prefix).
        // -------------------------------------------------
        // This is a 1×1 pixel transparent PNG encoded as Base64.
        string base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lK5XAAAAAElFTkSuQmCC";

        // -------------------------------------------------
        // 3. Create the data model for the report.
        // -------------------------------------------------
        ReportModel model = new ReportModel { ImageBase64 = base64Image };

        // -------------------------------------------------
        // 4. Build the report using the LINQ Reporting engine.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // -------------------------------------------------
        // 5. Save the generated report.
        // -------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
