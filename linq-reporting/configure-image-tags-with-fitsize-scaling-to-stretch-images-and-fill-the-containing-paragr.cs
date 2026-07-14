using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Path to the image that will be inserted into the report.
        public string ImagePath { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Directory.GetCurrentDirectory();

        // -----------------------------------------------------------------
        // 1. Create a sample image file (1x1 pixel PNG) that will be used
        //    by the image tag in the template.
        // -----------------------------------------------------------------
        string imageFile = Path.Combine(workDir, "sample.png");
        // Base64 encoded PNG (transparent 1x1 pixel).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        File.WriteAllBytes(imageFile, Convert.FromBase64String(base64Png));

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        //    The image tag is placed inside a textbox so that the
        //    -fitSize switch can stretch the image to fill the paragraph.
        // -----------------------------------------------------------------
        string templateFile = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will act as the image container.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting image tag with -fitSize switch.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template.
        templateDoc.Save(templateFile);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report using LINQ Reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templateFile);

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            ImagePath = imageFile
        };

        // Build the report. The root object name must match the tag prefix.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        string outputFile = Path.Combine(workDir, "Report.docx");
        reportDoc.Save(outputFile);
    }
}
