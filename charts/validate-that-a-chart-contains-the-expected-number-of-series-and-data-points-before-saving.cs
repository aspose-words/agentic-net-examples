using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Expected counts for validation.
        const int expectedSeriesCount = 2;
        const int expectedDataPointsPerSeries = 3;

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Define categories (X‑axis labels) and add two series with matching data points.
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        chart.Series.Add("Series 1", categories, new double[] { 10, 20, 30 });
        chart.Series.Add("Series 2", categories, new double[] { 15, 25, 35 });

        // ----- Validation -----
        // Verify the number of series.
        if (chart.Series.Count != expectedSeriesCount)
            throw new InvalidOperationException(
                $"Chart must contain {expectedSeriesCount} series, but found {chart.Series.Count}.");

        // Verify each series contains the expected number of data points.
        foreach (ChartSeries series in chart.Series)
        {
            if (series.DataPoints.Count != expectedDataPointsPerSeries)
                throw new InvalidOperationException(
                    $"Series '{series.Name}' must contain {expectedDataPointsPerSeries} data points, but found {series.DataPoints.Count}.");
        }

        // Save the validated document.
        doc.Save("validated-chart.docx");
    }
}
