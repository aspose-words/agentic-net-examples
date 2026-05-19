using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // Needed for the Shape class
using Aspose.Words.Drawing.Charts;

public class ChartValidationExample
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart and obtain its Shape object.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Define categories (X‑axis) and corresponding values (Y‑axis).
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 10.0, 20.0, 30.0, 40.0 };

        // Add a single series with the defined data.
        chart.Series.Add("Revenue", categories, values);

        // Expected counts.
        int expectedSeriesCount = 1;
        int expectedDataPointsPerSeries = categories.Length;

        // Validate the number of series.
        if (chart.Series.Count != expectedSeriesCount)
            throw new InvalidOperationException(
                $"Chart must contain {expectedSeriesCount} series, but found {chart.Series.Count}.");

        // Validate the number of data points in each series.
        foreach (ChartSeries series in chart.Series)
        {
            if (series.DataPoints.Count != expectedDataPointsPerSeries)
                throw new InvalidOperationException(
                    $"Series '{series.Name}' must contain {expectedDataPointsPerSeries} data points, but found {series.DataPoints.Count}.");
        }

        // Save the document.
        doc.Save("validated-chart.docx");
    }
}
