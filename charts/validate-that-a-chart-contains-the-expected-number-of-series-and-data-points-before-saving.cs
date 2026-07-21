using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Required for Shape
using Aspose.Words.Drawing.Charts;        // Chart APIs

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart. The default column chart contains 3 series,
        // each with 4 data points (categories).
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);

        // Ensure the shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Expected counts based on the default chart data.
        const int expectedSeriesCount = 3;
        const int expectedDataPointsPerSeries = 4;

        // Validate the number of series.
        if (chart.Series.Count != expectedSeriesCount)
            throw new InvalidOperationException(
                $"Chart series count mismatch. Expected: {expectedSeriesCount}, Actual: {chart.Series.Count}");

        // Validate each series contains the expected number of data points.
        for (int i = 0; i < chart.Series.Count; i++)
        {
            var series = chart.Series[i];
            if (series.DataPoints.Count != expectedDataPointsPerSeries)
                throw new InvalidOperationException(
                    $"Series '{series.Name}' data point count mismatch. Expected: {expectedDataPointsPerSeries}, Actual: {series.DataPoints.Count}");
        }

        // Save the validated document.
        doc.Save("validated-chart.docx");
    }
}
