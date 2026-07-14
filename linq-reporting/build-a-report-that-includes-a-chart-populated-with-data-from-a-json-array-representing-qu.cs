using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class QuarterlyResult
{
    public string Quarter { get; set; } = "";
    public double Revenue { get; set; }
}

public class ReportModel
{
    public List<QuarterlyResult> Results { get; set; } = new();
    public string ChartHtml { get; set; } = "";
    public string ReportDate { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");
}

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data.
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "quarterly.json");
        string jsonContent = @"[
            { ""Quarter"": ""Q1"", ""Revenue"": 120000 },
            { ""Quarter"": ""Q2"", ""Revenue"": 150000 },
            { ""Quarter"": ""Q3"", ""Revenue"": 130000 },
            { ""Quarter"": ""Q4"", ""Revenue"": 170000 }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // Load data from JSON.
        List<QuarterlyResult> results = JsonConvert.DeserializeObject<List<QuarterlyResult>>(jsonContent) ?? new();

        // Build simple HTML bar chart.
        string chartHtml = "<div style='font-family:Arial;'>";
        foreach (var r in results)
        {
            int barWidth = (int)(r.Revenue / 1000); // scale for display
            chartHtml += $"<div>{r.Quarter}: ${r.Revenue:N0}</div>";
            chartHtml += $"<div style='background:#4F81BD;height:15px;width:{barWidth}px;'></div>";
        }
        chartHtml += "</div>";

        // Prepare the model.
        ReportModel model = new()
        {
            Results = results,
            ChartHtml = chartHtml
        };

        // Create the template document.
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("Quarterly Financial Report");
        builder.Writeln($"Generated on: <<[model.ReportDate]>>");
        builder.Writeln();

        // Table of raw data.
        builder.Writeln("Quarterly Results:");
        builder.Writeln("<<foreach [r in model.Results]>>");
        builder.Writeln("<<[r.Quarter]>> - <<[r.Revenue]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Insert the HTML chart.
        builder.Writeln("<<html [model.ChartHtml]>>");

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(template, model, "model");

        // Save the report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "QuarterlyReport.docx");
        template.Save(outputPath);
    }
}
