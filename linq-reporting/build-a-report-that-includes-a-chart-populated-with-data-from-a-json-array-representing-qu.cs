using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class QuarterlyResult
{
    public string Quarter { get; set; } = "";
    public double Value { get; set; }
}

public class ReportModel
{
    public List<QuarterlyResult> Results { get; set; } = new();
    public string ChartHtml { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Sample JSON data representing quarterly results.
        string json = @"[
            {""Quarter"":""Q1"",""Value"":1200},
            {""Quarter"":""Q2"",""Value"":1500},
            {""Quarter"":""Q3"",""Value"":1100},
            {""Quarter"":""Q4"",""Value"":1700}
        ]";

        // Deserialize JSON into model objects.
        List<QuarterlyResult> results = JsonConvert.DeserializeObject<List<QuarterlyResult>>(json) ?? new();

        // Prepare the report model.
        ReportModel model = new ReportModel
        {
            Results = results,
            ChartHtml = GenerateChartHtml(results)
        };

        // Create the template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Quarterly Results Report");
        builder.Writeln();

        // Table of raw data.
        builder.Writeln("<<foreach [r in Results]>>");
        builder.Writeln("<<[r.Quarter]>>: <<[r.Value]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Insert the HTML chart.
        builder.Writeln("Chart:");
        builder.Writeln("<<html [ChartHtml]>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template (optional, can reuse the same document).
        Document doc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = "report.docx";
        doc.Save(reportPath);
    }

    private static string GenerateChartHtml(List<QuarterlyResult> results)
    {
        // Simple horizontal bar chart using inline CSS.
        const double scaleFactor = 0.1; // Adjust to fit the bar width.
        var rows = results.Select(r =>
        {
            int barWidth = (int)(r.Value * scaleFactor);
            string bar = $"<span style='display:inline-block;background-color:steelblue;height:20px;width:{barWidth}px'></span>";
            return $"<div>{r.Quarter}: {bar} {r.Value}</div>";
        });

        return $"<div style='font-family:Arial;'>{string.Join("", rows)}</div>";
    }
}
