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
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title placeholder using LINQ Reporting syntax.
        builder.Writeln("Report Title: <<[model.Title]>>");
        builder.Writeln();

        // Insert a column chart (Word 2016 chart) – will be populated later.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // Save the template (required by the lifecycle rule).
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // 2. Build a DataSet that holds numeric data for the chart.
        DataSet dataSet = new DataSet();
        DataTable valuesTable = new DataTable("Values");
        valuesTable.Columns.Add("Category", typeof(string));
        valuesTable.Columns.Add("Amount", typeof(double));

        valuesTable.Rows.Add("Q1", 1200.5);
        valuesTable.Rows.Add("Q2", 1500.0);
        valuesTable.Rows.Add("Q3", 1100.75);
        valuesTable.Rows.Add("Q4", 1800.25);

        dataSet.Tables.Add(valuesTable);

        // 3. Create a wrapper model that contains the title and the DataSet.
        ReportModel model = new ReportModel
        {
            Title = "Quarterly Sales",
            Data = dataSet
        };

        // 4. Load the template and run the LINQ Reporting engine.
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None // default options
        };
        engine.BuildReport(doc, model, "model");

        // 5. After the report is built, locate the chart and populate it with data from the DataSet.
        Shape chart = doc.GetChildNodes(NodeType.Shape, true)
                         .Cast<Shape>()
                         .FirstOrDefault(s => s.HasChart);

        if (chart != null && chart.HasChart)
        {
            Chart wordChart = chart.Chart;

            // Extract categories and values from the DataSet.
            string[] categories = model.Data.Tables["Values"]
                                      .AsEnumerable()
                                      .Select(r => r.Field<string>("Category"))
                                      .ToArray();

            double[] values = model.Data.Tables["Values"]
                                   .AsEnumerable()
                                   .Select(r => r.Field<double>("Amount"))
                                   .ToArray();

            // Clear any existing series and add a new one with the extracted data.
            wordChart.Series.Clear();
            wordChart.Series.Add("Sales", categories, values);
        }

        // 6. Save the final report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }
}

// Public data model required by the LINQ Reporting engine.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public DataSet Data { get; set; } = new DataSet();
}
