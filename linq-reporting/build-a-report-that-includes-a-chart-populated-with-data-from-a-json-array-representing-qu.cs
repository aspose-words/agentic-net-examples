using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Newtonsoft.Json;

public class QuarterlyResult
{
    public string Quarter { get; set; } = "";
    public double Revenue { get; set; }
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample JSON data.
        // -----------------------------------------------------------------
        string jsonPath = "quarterly.json";
        var sampleData = new List<QuarterlyResult>
        {
            new() { Quarter = "Q1", Revenue = 15000 },
            new() { Quarter = "Q2", Revenue = 18000 },
            new() { Quarter = "Q3", Revenue = 21000 },
            new() { Quarter = "Q4", Revenue = 24000 }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData));

        // -----------------------------------------------------------------
        // 2. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Quarterly Revenue Report");
        builder.Writeln();
        builder.Writeln("Data Table:");
        builder.Writeln("<<foreach [item in results]>>");
        builder.Writeln("<<[item.Quarter]>>\t<<[item.Revenue]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Insert a placeholder chart (will be populated after the report is built).
        builder.Writeln("Revenue Chart:");
        var chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        // Add a dummy series so the chart exists in the template.
        chartShape.Chart.Series.Add("Revenue", new[] { "Q1", "Q2", "Q3", "Q4" }, new[] { 0.0, 0.0, 0.0, 0.0 });

        // Save the template.
        string templatePath = "template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report using LINQ Reporting.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var jsonDataSource = new JsonDataSource(jsonPath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "results");

        // -----------------------------------------------------------------
        // 4. Populate the chart with real data.
        // -----------------------------------------------------------------
        var results = JsonConvert.DeserializeObject<List<QuarterlyResult>>(File.ReadAllText(jsonPath)) ?? new();

        // Locate the chart shape in the generated document.
        var shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        var chart = shape.Chart;

        // Replace the dummy series with actual data.
        chart.Series.Clear();
        var categories = new List<string>();
        var values = new List<double>();
        foreach (var item in results)
        {
            categories.Add(item.Quarter);
            values.Add(item.Revenue);
        }
        chart.Series.Add("Revenue", categories.ToArray(), values.ToArray());

        // -----------------------------------------------------------------
        // 5. Save the final report.
        // -----------------------------------------------------------------
        doc.Save("QuarterlyReport.docx");
    }
}
