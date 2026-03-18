using System;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added for Shape
using Aspose.Words.Drawing.Charts;   // Chart, ChartSeries, etc.

namespace ChartValidationExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart. By default Aspose.Words creates a chart with 3 series,
            // each containing 4 data points (one for each category).
            Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
            Chart chart = chartShape.Chart;

            // Expected numbers.
            const int expectedSeriesCount = 3;
            const int expectedDataPointsPerSeries = 4;

            // Validate the number of series.
            if (chart.Series.Count != expectedSeriesCount)
                throw new InvalidOperationException(
                    $"Chart contains {chart.Series.Count} series, but {expectedSeriesCount} were expected.");

            // Validate the number of data points in each series.
            foreach (ChartSeries series in chart.Series)
            {
                if (series.DataPoints.Count != expectedDataPointsPerSeries)
                    throw new InvalidOperationException(
                        $"Series \"{series.Name}\" contains {series.DataPoints.Count} data points, " +
                        $"but {expectedDataPointsPerSeries} were expected.");
            }

            // All validations passed – save the document.
            doc.Save("ValidatedChart.docx");
        }
    }
}
