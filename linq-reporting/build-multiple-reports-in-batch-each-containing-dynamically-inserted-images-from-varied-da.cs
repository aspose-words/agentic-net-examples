using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class ReportModel
{
    // Title displayed in the report.
    public string Title { get; set; } = string.Empty;

    // Base64-encoded PNG image (1x1 pixel) inserted into the report.
    public string ImageBase64 { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a reusable template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a placeholder for the report title.
        builder.Writeln("Report: <<[model.Title]>>");

        // Insert a textbox that will contain the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImageBase64] -fitSize>>");

        // Save the template to disk (required before building reports).
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare sample data for multiple reports.
        // -----------------------------------------------------------------
        // A 1x1 pixel transparent PNG encoded in Base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";

        List<ReportModel> reportData = new()
        {
            new ReportModel { Title = "First Report",  ImageBase64 = base64Png },
            new ReportModel { Title = "Second Report", ImageBase64 = base64Png },
            new ReportModel { Title = "Third Report",  ImageBase64 = base64Png }
        };

        // -----------------------------------------------------------------
        // 3. Build each report in batch using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        // Load the template once; clone it for each iteration to keep it unchanged.
        Document templateDoc = new Document(templatePath);

        int index = 1;
        foreach (ReportModel model in reportData)
        {
            // Clone the template for the current report.
            Document report = (Document)templateDoc.Clone();

            // Build the report by binding the model to the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(report, model, "model");

            // Save the generated report with a distinct filename.
            string outputFileName = $"Report_{index}_{model.Title.Replace(' ', '_')}.docx";
            report.Save(outputFileName);
            index++;
        }

        // -----------------------------------------------------------------
        // 4. Indicate completion.
        // -----------------------------------------------------------------
        Console.WriteLine("Batch report generation completed.");
    }
}
