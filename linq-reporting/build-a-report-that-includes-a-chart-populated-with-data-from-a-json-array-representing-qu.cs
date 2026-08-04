using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Newtonsoft.Json;

public class QuarterResult
{
    public string Quarter { get; set; } = "";
    public double Revenue { get; set; }
}

public class ReportModel
{
    public List<QuarterResult> Results { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "quarterly.json");
        var sampleData = new List<QuarterResult>
        {
            new() { Quarter = "Q1", Revenue = 12000 },
            new() { Quarter = "Q2", Revenue = 15000 },
            new() { Quarter = "Q3", Revenue = 13000 },
            new() { Quarter = "Q4", Revenue = 17000 }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load data into model
        var model = new ReportModel
        {
            Results = JsonConvert.DeserializeObject<List<QuarterResult>>(File.ReadAllText(jsonPath))!
        };

        // Create template document with LINQ Reporting tags
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Quarterly Revenue Report");
        builder.Writeln();

        // Table header
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Quarter");
        builder.InsertCell();
        builder.Writeln("Revenue");
        builder.EndRow();

        // Table rows populated from JSON via LINQ Reporting
        builder.Writeln("<<foreach [item in model.Results]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quarter]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Revenue]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save template
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load template for report generation
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Insert chart based on the same data
        var chartBuilder = new DocumentBuilder(reportDoc);
        chartBuilder.MoveToDocumentEnd();
        chartBuilder.Writeln();

        Shape chartShape = chartBuilder.InsertChart(ChartType.Column, 432, 288);
        Chart chart = chartShape.Chart;
        chart.Title.Text = "Quarterly Revenue";

        // Prepare data arrays for the series
        string[] categories = model.Results.Select(r => r.Quarter).ToArray();
        double[] values = model.Results.Select(r => r.Revenue).ToArray();

        // Clear default series and add our data
        chart.Series.Clear();
        chart.Series.Add("Revenue", categories, values);

        // Save final report
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "QuarterlyReport.docx");
        reportDoc.Save(outputPath);
    }
}
