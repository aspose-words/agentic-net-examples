using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // Needed for the Shape class
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a default column chart. The chart contains 3 demo series,
        // each with 4 data points (categories) by default.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Ensure the shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Expected counts based on the default demo data.
        const int expectedSeriesCount = 3;
        const int expectedDataPointsPerSeries = 4;

        // Validate the number of series.
        if (chart.Series.Count != expectedSeriesCount)
            throw new InvalidOperationException(
                $"Chart validation failed: expected {expectedSeriesCount} series, but found {chart.Series.Count}.");

        // Validate the number of data points in each series.
        foreach (ChartSeries series in chart.Series)
        {
            if (series.DataPoints.Count != expectedDataPointsPerSeries)
                throw new InvalidOperationException(
                    $"Chart validation failed: series \"{series.Name}\" expected {expectedDataPointsPerSeries} data points, but found {series.DataPoints.Count}.");
        }

        // All validations passed – save the document.
        doc.Save("validated-chart.docx");
    }
}
