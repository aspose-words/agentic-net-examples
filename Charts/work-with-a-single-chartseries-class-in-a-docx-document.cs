using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing; // Added for Shape
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a width of 500 points and a height of 300 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series that Aspose.Words adds to a new chart.
        chart.Series.Clear();

        // Define the categories (X‑axis) and the corresponding values (Y‑axis) for a single series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120.5, 150.0, 130.2, 170.8 };

        // Add a single series to the chart.
        // This uses the overload: Add(string seriesName, string[] categories, double[] values)
        ChartSeries series = chart.Series.Add("Revenue", categories, values);

        // Set the fill color of the series to green.
        series.Format.Fill.ForeColor = Color.Green;

        // Enable data labels for the series and configure them to show the value.
        series.HasDataLabels = true;
        // Apply the same label settings to all data points in the series.
        for (int i = 0; i < series.DataLabels.Count; i++)
        {
            series.DataLabels[i].ShowValue = true;
            series.DataLabels[i].ShowCategoryName = false;
            series.DataLabels[i].ShowSeriesName = false;
        }

        // Save the document to a DOCX file.
        doc.Save("ChartSeriesExample.docx");
    }
}
