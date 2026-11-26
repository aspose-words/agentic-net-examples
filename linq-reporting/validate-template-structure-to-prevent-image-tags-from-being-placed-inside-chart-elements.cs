using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for template and result documents.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string resultPath = Path.Combine(outputDir, "report.docx");

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a chart shape (the element we want to protect).
        builder.InsertChart(ChartType.Column, 400, 300);

        // Insert a simple placeholder outside the chart.
        builder.Writeln("<<[model.Title]>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for validation.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Validate that no image tags are placed inside chart shapes.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Chart shapes expose a Chart object; check for its presence.
            if (shape.HasChart)
            {
                // Get any text inside the chart shape.
                string innerText = shape.GetText();

                if (!string.IsNullOrEmpty(innerText) && innerText.Contains("<<image"))
                {
                    Console.WriteLine("Invalid template: image tag found inside a chart element.");
                    return;
                }
            }
        }

        // -----------------------------------------------------------------
        // Build the report using LINQ Reporting Engine.
        // -----------------------------------------------------------------
        ReportModel model = new()
        {
            Title = "Sales Overview"
        };

        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(resultPath);
        Console.WriteLine($"Report generated successfully at: {resultPath}");
    }
}

// Simple data model aligned with the template.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
}
