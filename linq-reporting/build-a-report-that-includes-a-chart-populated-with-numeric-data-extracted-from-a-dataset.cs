using System;
using System.Data;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ReportModel
{
    // Title displayed in the report.
    public string Title { get; set; } = "Sales Report";
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample data in a DataSet.
        // -----------------------------------------------------------------
        DataSet dataSet = new DataSet();
        DataTable salesTable = new DataTable("Sales");
        salesTable.Columns.Add("Month", typeof(string));
        salesTable.Columns.Add("Amount", typeof(double));

        salesTable.Rows.Add("Jan", 1200.5);
        salesTable.Rows.Add("Feb", 1500.0);
        salesTable.Rows.Add("Mar", 1800.75);
        salesTable.Rows.Add("Apr", 1100.25);
        salesTable.Rows.Add("May", 1700.0);
        dataSet.Tables.Add(salesTable);

        // -----------------------------------------------------------------
        // 2. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a placeholder for the report title using LINQ Reporting syntax.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Insert an empty chart that will be populated later.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        // The chart initially contains a default series; it will be replaced after the report is built.

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportModel model = new ReportModel();

        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        // Populate the title placeholder.
        engine.BuildReport(report, model, "model");

        // -----------------------------------------------------------------
        // 4. Populate the chart with data from the DataSet.
        // -----------------------------------------------------------------
        // Locate the chart shape inside the document.
        Shape chartContainer = (Shape)report.GetChildNodes(NodeType.Shape, true)
                                            .Cast<Node>()
                                            .FirstOrDefault(s => ((Shape)s).HasChart);

        if (chartContainer != null && chartContainer.HasChart)
        {
            Chart chart = chartContainer.Chart;

            // Remove any existing series.
            chart.Series.Clear();

            // Prepare categories (X‑axis) and values (Y‑axis) from the DataTable.
            string[] categories = salesTable.AsEnumerable()
                                            .Select(row => row.Field<string>("Month"))
                                            .ToArray();

            double[] values = salesTable.AsEnumerable()
                                        .Select(row => row.Field<double>("Amount"))
                                        .ToArray();

            // Add a new series with the extracted data.
            chart.Series.Add("Sales", categories, values);

            // Optional: set a chart title.
            chart.Title.Text = "Monthly Sales";
        }

        // -----------------------------------------------------------------
        // 5. Save the final report.
        // -----------------------------------------------------------------
        const string reportPath = "Report.docx";
        report.Save(reportPath);
    }
}
