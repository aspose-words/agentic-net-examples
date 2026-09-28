using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class QuarterResult
{
    public string Quarter { get; set; } = "";
    public double Result { get; set; }
}

public class ReportModel
{
    public List<QuarterResult> Results { get; set; } = new();
    public string ChartHtml { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 1. Create sample JSON data
        var sampleData = new List<QuarterResult>
        {
            new() { Quarter = "Q1", Result = 120 },
            new() { Quarter = "Q2", Result = 150 },
            new() { Quarter = "Q3", Result = 90 },
            new() { Quarter = "Q4", Result = 180 }
        };
        string jsonPath = "data.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // 2. Load JSON data into model
        var results = JsonConvert.DeserializeObject<List<QuarterResult>>(File.ReadAllText(jsonPath)) ?? new();
        var model = new ReportModel { Results = results };
        model.ChartHtml = BuildChartHtml(model.Results);

        // 3. Create template document with LINQ Reporting tag
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Quarterly Results Report");
        builder.Writeln("<<html [model.ChartHtml]>>");

        // Save template (optional, for inspection)
        string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // 4. Build the report using ReportingEngine
        var engine = new ReportingEngine();
        bool success = engine.BuildReport(templateDoc, model, "model");

        // 5. Save the generated report
        string reportPath = "Report.docx";
        templateDoc.Save(reportPath);

        // Output simple status (no interactive input)
        Console.WriteLine(success
            ? $"Report generated successfully: {reportPath}"
            : "Report generation failed.");
    }

    private static string BuildChartHtml(List<QuarterResult> results)
    {
        var sb = new StringBuilder();
        sb.Append("<div style='font-family:Arial;'><h2>Quarterly Results</h2>");
        double max = 0;
        foreach (var r in results)
        {
            if (r.Result > max) max = r.Result;
        }
        // Scale factor to keep bars within reasonable width
        double scale = max > 0 ? 300 / max : 1;

        foreach (var r in results)
        {
            double width = r.Result * scale;
            sb.Append($"<div style='margin:5px 0;'>{r.Quarter}: ");
            sb.Append($"<span style='display:inline-block;background:#4CAF50;height:20px;width:{width}px;'></span> ");
            sb.Append($"{r.Result}</div>");
        }

        sb.Append("</div>");
        return sb.ToString();
    }
}
