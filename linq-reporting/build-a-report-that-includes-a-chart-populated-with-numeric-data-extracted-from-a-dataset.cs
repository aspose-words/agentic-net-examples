using System;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "ReportWithChart.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a title that will be filled by the reporting engine.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln(); // empty line.

        // Insert a placeholder chart (Column chart). The chart will be populated later.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        // The chart initially contains a dummy series; it will be replaced after BuildReport.

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data source (DataSet with a numeric table).
        // -----------------------------------------------------------------
        DataSet dataSet = new DataSet();

        DataTable valuesTable = new DataTable("Values");
        valuesTable.Columns.Add("Category", typeof(string));
        valuesTable.Columns.Add("Amount", typeof(double));

        valuesTable.Rows.Add("Q1", 1200.5);
        valuesTable.Rows.Add("Q2", 1500.0);
        valuesTable.Rows.Add("Q3", 1100.75);
        valuesTable.Rows.Add("Q4", 1800.25);

        dataSet.Tables.Add(valuesTable);

        // -----------------------------------------------------------------
        // 3. Create the root model object for the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Title = "Quarterly Sales Report",
            Data = dataSet
        };

        // -----------------------------------------------------------------
        // 4. Load the template and build the report using ReportingEngine.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Populate the chart with data from the DataSet table.
        // -----------------------------------------------------------------
        // Locate the first chart shape in the document.
        Shape? chartContainer = reportDoc.GetChildNodes(NodeType.Shape, true)
                                         .Cast<Shape>()
                                         .FirstOrDefault(s => s.HasChart);

        if (chartContainer != null && chartContainer.HasChart)
        {
            Chart chart = chartContainer.Chart;

            // Remove any existing series.
            chart.Series.Clear();

            // Extract categories and values from the DataTable.
            string[] categories = valuesTable.AsEnumerable()
                                             .Select(row => row.Field<string>("Category"))
                                             .ToArray();

            double[] values = valuesTable.AsEnumerable()
                                         .Select(row => row.Field<double>("Amount"))
                                         .ToArray();

            // Add a new series with the extracted data.
            chart.Series.Add("Sales", categories, values);
        }

        // -----------------------------------------------------------------
        // 6. Save the final report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }

    // Root model class used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Title displayed in the report.
        public string Title { get; set; } = string.Empty;

        // DataSet containing the numeric data for the chart.
        public DataSet Data { get; set; } = new DataSet();
    }
}
