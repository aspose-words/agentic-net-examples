using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class QuarterlyResult
{
    public string Quarter { get; set; } = "";
    public double Revenue { get; set; }
}

public class ReportModel
{
    public List<QuarterlyResult> Results { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample JSON data.
        string json = @"[
            { ""Quarter"": ""Q1"", ""Revenue"": 12000 },
            { ""Quarter"": ""Q2"", ""Revenue"": 15000 },
            { ""Quarter"": ""Q3"", ""Revenue"": 13000 },
            { ""Quarter"": ""Q4"", ""Revenue"": 17000 }
        ]";

        // Deserialize JSON.
        ReportModel model = new()
        {
            Results = JsonConvert.DeserializeObject<List<QuarterlyResult>>(json) ?? new List<QuarterlyResult>()
        };

        // Create template document.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Quarterly Results Report");
        builder.Writeln();

        // Table with foreach.
        builder.Writeln("<<foreach [r in Results]>>");
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Quarter");
        builder.InsertCell();
        builder.Writeln("Revenue");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[r.Quarter]>>");
        builder.InsertCell();
        builder.Writeln("<<[r.Revenue]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        builder.Writeln();
        builder.Writeln("Chart of Quarterly Revenue:");
        // Bookmark placeholder for chart.
        builder.Writeln("<<bookmark [\"ChartPlaceholder\"]>>Chart will be inserted here<</bookmark>>");

        // Save template.
        templateDoc.Save(templatePath);

        // Load template and build report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Insert chart at bookmark.
        Bookmark chartBookmark = reportDoc.Range.Bookmarks["ChartPlaceholder"];
        if (chartBookmark != null)
        {
            DocumentBuilder chartBuilder = new DocumentBuilder(reportDoc);
            chartBuilder.MoveToBookmark(chartBookmark.Name);
            chartBuilder.Writeln(); // New paragraph.

            // Insert chart.
            Shape chartShape = chartBuilder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = chartShape.Chart;

            chart.Title.Text = "Quarterly Revenue";

            // Clear default series and add our data.
            chart.Series.Clear();
            ChartSeries series = chart.Series.Add(
                "Revenue",
                model.Results.Select(r => r.Quarter).ToArray(),
                model.Results.Select(r => r.Revenue).ToArray());

            // Axis titles.
            chart.AxisX.Title.Text = "Quarter";
            chart.AxisY.Title.Text = "Revenue";
        }

        // Save final report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "QuarterlyReport.docx");
        reportDoc.Save(outputPath);
    }
}
