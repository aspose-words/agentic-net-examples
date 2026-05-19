using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Prepare a sample image as a byte array (1x1 transparent PNG).
        // -----------------------------------------------------------------
        // Base64 representation of a minimal PNG image.
        const string pngBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2ZcAAAAASUVORK5CYII=";
        byte[] sampleImageBytes = Convert.FromBase64String(pngBase64);

        // -----------------------------------------------------------------
        // 2. Build the LINQ Reporting template programmatically.
        //    The image tag is placed inside a textbox and uses the -fitWidth switch.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 400, 200);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting tag: pass the image byte array and limit its width.
        builder.Write("<<image [model.ImageData] -fitWidth>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var model = new ReportModel { ImageData = sampleImageBytes };

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the final report.
        // -----------------------------------------------------------------
        reportDoc.Save(reportPath);
    }
}

// Data model exposed to the template. The ImageData property supplies the image bytes.
public class ReportModel
{
    // Initialized via object initializer to avoid nullable warnings.
    public byte[] ImageData { get; set; } = null!;
}
