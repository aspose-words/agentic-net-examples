using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";
        const string sampleImagePath = "SampleImage.png";

        // Ensure a sample image exists for the image tag.
        CreateSampleImage(sampleImagePath);

        // 1. Create a template document that contains a chart and a valid image tag.
        CreateTemplate(templatePath);

        // 2. Validate that no image tags are placed inside chart elements.
        bool templateIsValid = ValidateTemplate(templatePath);
        Console.WriteLine(templateIsValid
            ? "Template validation passed."
            : "Template validation failed: image tag found inside a chart.");

        // 3. If the template is valid, build the report using LINQ Reporting.
        if (templateIsValid)
        {
            // Sample data model.
            var model = new ReportModel
            {
                Title = "Quarterly Sales Report",
                ImagePath = sampleImagePath,
                ChartData = new()
                {
                    new ChartItem { Category = "Q1", Value = 1200 },
                    new ChartItem { Category = "Q2", Value = 1500 },
                    new ChartItem { Category = "Q3", Value = 1800 },
                    new ChartItem { Category = "Q4", Value = 2000 }
                }
            };

            // Load the template, populate it, and save the final document.
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");
            doc.Save(reportPath);
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }

    // Creates a simple 1x1 PNG file to be used by the image tag.
    private static void CreateSampleImage(string path)
    {
        // Base64-encoded 1x1 transparent PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lK6XAAAAAElFTkSuQmCC";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, pngBytes);
    }

    // Creates a template with a chart and a correctly placed image tag (inside a textbox).
    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title placeholder.
        builder.Writeln("<<[model.Title]>>");

        // Insert a chart shape (no image tags inside).
        builder.InsertChart(ChartType.Column, 400, 300);

        // Insert a textbox that will hold the image tag (valid placement).
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template.
        doc.Save(path);
    }

    // Checks all chart shapes for the presence of an image tag.
    private static bool ValidateTemplate(string path)
    {
        var doc = new Document(path);

        // Retrieve all shapes that contain a chart.
        var chartShapes = doc.GetChildNodes(NodeType.Shape, true)
            .Cast<Shape>()
            .Where(s => s.HasChart);

        // If any chart shape's inner text contains an image tag, the template is invalid.
        foreach (var chart in chartShapes)
        {
            string innerText = chart.GetText();
            if (innerText.Contains("<<image"))
                return false;
        }

        return true;
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    public string Title { get; set; } = "Untitled Report";
    public string ImagePath { get; set; } = string.Empty;
    public List<ChartItem> ChartData { get; set; } = new();
}

// Simple class representing a data point for the chart.
public class ChartItem
{
    public string Category { get; set; } = string.Empty;
    public double Value { get; set; }
}
