using System;
using System.Data;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // -------------------- Prepare sample data --------------------
        DataSet dataSet = new DataSet();
        DataTable table = new DataTable("Data");
        table.Columns.Add("Month", typeof(string));
        table.Columns.Add("Value", typeof(double));

        table.Rows.Add("Jan", 1200.5);
        table.Rows.Add("Feb", 1500.0);
        table.Rows.Add("Mar", 1700.75);
        table.Rows.Add("Apr", 1600.25);
        table.Rows.Add("May", 1800.0);
        dataSet.Tables.Add(table);

        // -------------------- Create a template document --------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title.
        builder.Writeln("Sales Report");
        builder.Writeln();

        // Table template using LINQ Reporting tags.
        // Iterate over the DataTable named "Data".
        builder.Writeln("<<foreach [row in Data]>>");
        builder.Writeln("<<[row.Month]>>\t<<[row.Value]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Save the template (required by the workflow).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -------------------- Build the report --------------------
        // Load the template.
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Pass the DataTable itself as the data source; name it "Data" to match the tags.
        DataTable dataTable = dataSet.Tables["Data"];
        engine.BuildReport(report, dataTable, "Data");

        // -------------------- Insert a chart and populate it with the same data --------------------
        DocumentBuilder chartBuilder = new DocumentBuilder(report);
        chartBuilder.MoveToDocumentEnd();

        // Insert a column chart.
        Shape chartShape = chartBuilder.InsertChart(ChartType.Column, 432, 288);
        Chart chart = chartShape.Chart;

        // Extract categories and values from the DataTable.
        string[] categories = dataTable.AsEnumerable()
                                       .Select(r => r.Field<string>("Month"))
                                       .ToArray();

        double[] values = dataTable.AsEnumerable()
                                   .Select(r => r.Field<double>("Value"))
                                   .ToArray();

        // Clear any default series and add our data.
        chart.Series.Clear();
        chart.Series.Add("Sales", categories, values);

        // Save the final report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
