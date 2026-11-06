using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Working directory and file paths.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "output.docx");
        string imagePath = Path.Combine(workDir, "sample.png");

        // Create a simple 1x1 red PNG image for the report.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // -----------------------------------------------------------------
        // Step 1: Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a chart shape.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        // Intentionally place an image tag inside the chart title to simulate an invalid template.
        chartShape.Chart.Title.Text = "<<image [model.ImagePath]>>";

        // Add a textbox containing a valid image tag outside the chart.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template for validation.
        // -----------------------------------------------------------------
        var templateDoc = new Document(templatePath);

        // Validate that no image tags are placed inside chart elements.
        bool invalidFound = false;
        foreach (Shape shape in templateDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasChart)
            {
                string title = shape.Chart.Title.Text ?? string.Empty;
                if (title.Contains("<<image"))
                {
                    invalidFound = true;
                    Console.WriteLine("Validation error: Image tag found inside a chart title.");
                    break;
                }
            }
        }

        if (invalidFound)
        {
            // Abort report generation due to invalid template structure.
            Console.WriteLine("Report generation aborted due to template validation failure.");
            return;
        }

        // -----------------------------------------------------------------
        // Step 3: Build the report (template is valid).
        // -----------------------------------------------------------------
        var model = new ReportModel { ImagePath = imagePath };
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        templateDoc.Save(outputPath);
        Console.WriteLine($"Report generated successfully: {outputPath}");
    }
}

// Simple data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Path to the image file used in the template.
    public string ImagePath { get; set; } = string.Empty;
}
