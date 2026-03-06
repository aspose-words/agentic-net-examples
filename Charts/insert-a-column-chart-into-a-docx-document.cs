using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added for Shape
using Aspose.Words.Drawing.Charts;   // Chart, ChartType

class InsertColumnChartExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart shape with the desired size (width and height are in points).
        // ChartType.Column specifies a 2‑D column chart.
        double chartWidth = 400;   // points
        double chartHeight = 300;  // points
        Shape chartShape = builder.InsertChart(ChartType.Column, chartWidth, chartHeight);

        // Get the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Define categories (X‑axis labels) and corresponding values (Y‑axis).
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120.5, 150.0, 130.75, 170.25 };

        // Add a series to the chart. For column charts the Add method takes
        // a series name, an array of category names and an array of numeric values.
        chart.Series.Add("Sales", categories, values);

        // Optionally set a title for the chart.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // Save the document to a DOCX file.
        doc.Save("ColumnChart.docx");
    }
}
