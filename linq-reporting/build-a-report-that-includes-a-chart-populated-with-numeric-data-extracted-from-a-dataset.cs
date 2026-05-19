using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing.Charts;

public class ReportModel
{
    // Collection of chart series that will be bound to the chart tag.
    public List<ChartSeries> Chart { get; set; } = new();
}

// Simple POCO representing a chart series for LINQ Reporting.
public class ChartSeries
{
    public string Name { get; set; } = string.Empty;
    public List<double> Values { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- Create sample data in a DataSet ----------
        var dataSet = new DataSet();
        var table = new DataTable("Sales");
        table.Columns.Add("Month", typeof(string));
        table.Columns.Add("Revenue", typeof(double));

        table.Rows.Add("Jan", 12000.5);
        table.Rows.Add("Feb", 15000.75);
        table.Rows.Add("Mar", 17000);
        table.Rows.Add("Apr", 13000.25);
        dataSet.Tables.Add(table);

        // ---------- Convert DataSet data to a model suitable for LINQ Reporting ----------
        var model = new ReportModel();

        var series = new ChartSeries
        {
            Name = "Revenue",
            // Extract numeric values from the DataTable.
            Values = table.AsEnumerable()
                         .Select(row => row.Field<double>("Revenue"))
                         .ToList()
        };

        model.Chart.Add(series);

        // ---------- Build the template document programmatically ----------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Quarterly Revenue Chart");

        // Insert a chart shape and populate it with the data extracted from the DataSet.
        var chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        var chart = chartShape.Chart;

        // Clear any default series and add our own.
        chart.Series.Clear();

        // Use the first (and only) series from the model.
        var modelSeries = model.Chart.First();

        // Create X‑axis categories (e.g., 1, 2, 3, …) and add Y‑values from the model.
        for (int i = 0; i < modelSeries.Values.Count; i++)
        {
            double yValue = modelSeries.Values[i];
            // X value is the index (starting from 1) to keep the chart simple.
            chart.Series.Add(modelSeries.Name,
                             new[] { (double)(i + 1) },
                             new[] { yValue });
        }

        // ---------- Build the report ----------
        var engine = new Aspose.Words.Reporting.ReportingEngine();
        engine.Options = Aspose.Words.Reporting.ReportBuildOptions.None; // default options
        engine.BuildReport(doc, model, "model");

        // ---------- Save the generated report ----------
        doc.Save("ReportWithChart.docx");
    }
}
