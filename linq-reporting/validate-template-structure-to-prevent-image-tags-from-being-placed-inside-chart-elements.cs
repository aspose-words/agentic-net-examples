using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Path to a sample image file.
        public string ImagePath { get; set; } = "sample.png";
    }

    public static void Main()
    {
        // Ensure the working directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample image file (a tiny PNG).
        string imagePath = Path.Combine(outputDir, "sample.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAn8B9pXK3VQAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create the template document.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report Header");

        // Insert a chart and deliberately place an image tag inside its title (invalid scenario).
        Chart chart = builder.InsertChart(ChartType.Column, 400, 300).Chart;
        chart.Title.Text = "<<image [model.ImagePath]>>";

        builder.Writeln("Report Footer");
        templateDoc.Save(templatePath);

        // Load the template for validation.
        Document loadedDoc = new Document(templatePath);

        bool hasInvalidImageTag = false;

        // Iterate over all chart objects in the document.
        var chartShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                   .OfType<Shape>()
                                   .Where(s => s.HasChart);

        foreach (Shape shape in chartShapes)
        {
            Chart c = shape.Chart;
            if (c.Title != null && !string.IsNullOrEmpty(c.Title.Text) &&
                c.Title.Text.Contains("<<image"))
            {
                hasInvalidImageTag = true;
                Console.WriteLine("Validation Error: Image tag found inside a chart title.");
            }
        }

        if (hasInvalidImageTag)
        {
            // Abort report generation due to invalid template structure.
            Console.WriteLine("Report generation aborted because the template contains invalid image tags inside chart elements.");
            return;
        }

        // If validation passes, build the report using LINQ Reporting.
        ReportModel model = new ReportModel { ImagePath = imagePath };
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);
        Console.WriteLine($"Report generated successfully: {reportPath}");
    }
}
