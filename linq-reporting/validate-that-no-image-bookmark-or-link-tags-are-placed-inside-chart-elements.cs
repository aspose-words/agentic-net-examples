using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Title { get; set; } = "Sample Report";
    public string ImagePath { get; set; } = "";
    public string Url { get; set; } = "https://example.com";
    public string LinkText { get; set; } = "Example Link";
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Prepare a simple 1x1 PNG image.
        // -----------------------------------------------------------------
        const string imageFileName = "sample.png";
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imageFileName, pngBytes);

        // -----------------------------------------------------------------
        // Create the data model.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            ImagePath = Path.GetFullPath(imageFileName)
        };

        // -----------------------------------------------------------------
        // Create the template document.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Simple text with a data field.
        builder.Writeln("Report Title: <<[model.Title]>>");

        // Insert a chart (no tags inside).
        builder.Writeln("Chart:");
        builder.InsertChart(ChartType.Column, 400, 300);

        // Valid tags placed outside the chart.
        builder.Writeln("<<bookmark [model.Title]>>Bookmark Content<</bookmark>>");
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");

        // Image tag must be inside a textbox (image container).
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for validation and reporting.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // Validate that no image, bookmark, or link tags exist inside chart elements.
        bool invalidTagFound = false;
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasChart)
            {
                string chartText = shape.GetText();
                if (chartText.Contains("<<image") ||
                    chartText.Contains("<<bookmark") ||
                    chartText.Contains("<<link"))
                {
                    invalidTagFound = true;
                    break;
                }
            }
        }

        Console.WriteLine(invalidTagFound
            ? "Invalid tags were found inside a chart."
            : "No invalid tags inside charts were detected.");

        // -----------------------------------------------------------------
        // Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
