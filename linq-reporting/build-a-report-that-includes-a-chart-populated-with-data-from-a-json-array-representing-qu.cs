using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // 1. Sample JSON data for quarterly results.
        string json = @"
        [
            { ""Quarter"": ""Q1"", ""Revenue"": 12000 },
            { ""Quarter"": ""Q2"", ""Revenue"": 15000 },
            { ""Quarter"": ""Q3"", ""Revenue"": 13000 },
            { ""Quarter"": ""Q4"", ""Revenue"": 17000 }
        ]";

        // 2. Deserialize JSON into a list of results.
        List<QuarterResult> results = JsonConvert.DeserializeObject<List<QuarterResult>>(json) ?? new();

        // 3. Generate a placeholder chart image (PNG byte array).
        byte[] chartImage = GeneratePlaceholderChartImage();

        // 4. Prepare the model for the LINQ Reporting engine.
        ReportModel model = new()
        {
            Results = results,
            ChartImage = chartImage
        };

        // 5. Create a Word template programmatically.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // 6. Load the template and build the report.
        Document doc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // 7. Save the final report.
        string reportPath = "Report.docx";
        doc.Save(reportPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }

    // Creates a simple template containing an image placeholder inside a textbox and a foreach table.
    private static void CreateTemplate(string path)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        builder.Writeln("Quarterly Revenue Report");
        builder.Writeln();

        // Image placeholder inside a textbox (required by LINQ Reporting).
        Shape chartBox = builder.InsertShape(ShapeType.TextBox, 500, 300);
        builder.MoveTo(chartBox.FirstParagraph);
        builder.Write("<<image [ChartImage] -fitSize>>");
        builder.Writeln();

        // Table header.
        builder.Writeln("Quarterly Data:");
        builder.Writeln();

        // Foreach loop to list each quarter's revenue.
        builder.Writeln("<<foreach [r in Results]>>");
        builder.Writeln("<<[r.Quarter]>>: $<<[r.Revenue]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }

    // Returns a simple 1×1 pixel PNG image as a byte array (placeholder for the chart).
    private static byte[] GeneratePlaceholderChartImage()
    {
        // Base64-encoded PNG (1×1 pixel, solid blue).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        return Convert.FromBase64String(base64Png);
    }
}

// Model representing a single quarter's result.
public class QuarterResult
{
    public string Quarter { get; set; } = "";
    public double Revenue { get; set; }
}

// Root model passed to the ReportingEngine.
public class ReportModel
{
    public List<QuarterResult> Results { get; set; } = new();
    public byte[] ChartImage { get; set; } = Array.Empty<byte>();
}
