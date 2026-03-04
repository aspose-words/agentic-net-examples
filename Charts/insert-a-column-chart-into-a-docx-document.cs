using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added for Shape
using Aspose.Words.Drawing.Charts;   // Chart related types

class InsertColumnChartExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a width of 400 points and a height of 300 points.
        // The InsertChart method returns a Shape that contains the Chart object.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series that Aspose.Words inserts.
        chart.Series.Clear();

        // Define categories (X‑axis labels) for the column chart.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };

        // Add two series of values, each series will be displayed as a set of columns.
        chart.Series.Add("Revenue", categories, new double[] { 15000, 21000, 18000, 24000 });
        chart.Series.Add("Expenses", categories, new double[] { 12000, 16000, 13000, 19000 });

        // Optionally set a title for the chart.
        chart.Title.Text = "Quarterly Financial Overview";
        chart.Title.Show = true;

        // Save the document in DOCX format.
        doc.Save("ColumnChart.docx");
    }
}
