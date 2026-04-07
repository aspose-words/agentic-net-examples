using System;
using System.Data;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class LinqReportingChartExample
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings that might be used.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Prepare sample data in a DataSet with a DataTable named "Sales".
        // -----------------------------------------------------------------
        DataSet dataSet = new DataSet("ReportData");
        DataTable salesTable = new DataTable("Sales");
        salesTable.Columns.Add("Month", typeof(string));
        salesTable.Columns.Add("Amount", typeof(double));

        salesTable.Rows.Add("Jan", 1200.5);
        salesTable.Rows.Add("Feb", 1500.0);
        salesTable.Rows.Add("Mar", 1100.75);
        salesTable.Rows.Add("Apr", 1700.25);
        salesTable.Rows.Add("May", 1600.0);
        dataSet.Tables.Add(salesTable);

        // ---------------------------------------------------------------
        // 2. Create a template document programmatically and save it.
        // ---------------------------------------------------------------
        const string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Sales Report");
        // Register DateTime as a known type so the expression <<[DateTime.Now]>> works.
        builder.Writeln("Generated on: <<[DateTime.Now]>>");
        builder.Writeln(); // empty line

        // Insert a placeholder chart (will be populated later).
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        // Store the shape name so we can retrieve it after the report is built.
        string chartShapeName = chartShape.Name;

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Load the template and run the LINQ Reporting engine.
        // ---------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        // Allow the engine to resolve static members like DateTime.Now.
        engine.KnownTypes.Add(typeof(DateTime));

        // Build the report using the DataSet as the data source.
        engine.BuildReport(reportDoc, dataSet, "data");

        // ---------------------------------------------------------------
        // 4. After the report is built, locate the chart and fill it with data.
        // ---------------------------------------------------------------
        // Find the chart shape by its stored name.
        Shape foundChartShape = reportDoc.GetChildNodes(NodeType.Shape, true)
                                         .Cast<Shape>()
                                         .FirstOrDefault(s => s.Name == chartShapeName);

        if (foundChartShape != null && foundChartShape.HasChart)
        {
            Chart chart = foundChartShape.Chart;
            chart.Title.Text = "Monthly Sales";

            // Clear any existing series.
            chart.Series.Clear();

            // Prepare category (Month) and value (Amount) arrays.
            string[] categories = salesTable.Rows
                                            .Cast<DataRow>()
                                            .Select(r => r["Month"].ToString())
                                            .ToArray();

            double[] values = salesTable.Rows
                                        .Cast<DataRow>()
                                        .Select(r => Convert.ToDouble(r["Amount"]))
                                        .ToArray();

            // Add a new series with the prepared data.
            chart.Series.Add("Sales", categories, values);
        }

        // ---------------------------------------------------------------
        // 5. Save the final report.
        // ---------------------------------------------------------------
        const string reportPath = "SalesReport.docx";
        reportDoc.Save(reportPath);
    }
}
