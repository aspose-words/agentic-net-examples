using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing.Charts;

public class SaleItem
{
    public string Month { get; set; } = "";
    public double Amount { get; set; }
}

public class ReportModel
{
    // Collection of sales items used by the LINQ Reporting template.
    public List<SaleItem> Sales { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for older encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data in a DataSet.
        var dataSet = new DataSet();
        var salesTable = new DataTable("Sales");
        salesTable.Columns.Add("Month", typeof(string));
        salesTable.Columns.Add("Amount", typeof(double));
        salesTable.Rows.Add("January", 12000.5);
        salesTable.Rows.Add("February", 15000.0);
        salesTable.Rows.Add("March", 17000.75);
        salesTable.Rows.Add("April", 13000.25);
        dataSet.Tables.Add(salesTable);

        // Populate the model with strongly‑typed items.
        var model = new ReportModel();
        foreach (DataRow row in salesTable.Rows)
        {
            model.Sales.Add(new SaleItem
            {
                Month = row.Field<string>("Month") ?? "",
                Amount = row.Field<double>("Amount")
            });
        }

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Sales Report");
        builder.Writeln();

        // Table header and data rows using LINQ Reporting tags.
        builder.Writeln("<<foreach [sale in Sales]>>");
        builder.Writeln("Month: <<[sale.Month]>>, Amount: <<[sale.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for report generation.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Build the report using LINQ Reporting.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // Insert a chart and populate it with the same data.
        // -----------------------------------------------------------------
        var chartBuilder = new DocumentBuilder(reportDoc);
        chartBuilder.MoveToDocumentEnd();
        var chart = chartBuilder.InsertChart(ChartType.Column, 500, 300).Chart;

        // Prepare data for the chart.
        string[] categories = model.Sales
            .Select(s => s.Month)
            .ToArray();

        double[] values = model.Sales
            .Select(s => s.Amount)
            .ToArray();

        // Clear any default series and add our data.
        chart.Series.Clear();
        var series = chart.Series.Add("Sales", categories, values);
        series.Name = "Monthly Sales";

        // Save the final report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
