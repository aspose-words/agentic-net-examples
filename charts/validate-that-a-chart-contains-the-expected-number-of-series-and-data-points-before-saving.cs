using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);

        // Ensure the shape actually contains a chart.
        if (!chartShape.HasChart)
        {
            throw new InvalidOperationException("Inserted shape does not contain a chart.");
        }

        // Access the chart.
        Chart chart = chartShape.Chart;

        // Clear any default series.
        chart.Series.Clear();

        // Add series with data points.
        chart.Series.Add("Series 1", new double[] { 10, 20, 30 });
        chart.Series.Add("Series 2", new double[] { 15, 25, 35 });

        // Expected counts.
        int expectedSeriesCount = 2;
        int expectedPointsPerSeries = 3;

        // Validate series count.
        if (chart.Series.Count != expectedSeriesCount)
        {
            throw new InvalidOperationException($"Chart must contain {expectedSeriesCount} series, but found {chart.Series.Count}.");
        }

        // Validate data point count for each series.
        foreach (ChartSeries series in chart.Series)
        {
            if (series.DataPoints.Count != expectedPointsPerSeries)
            {
                throw new InvalidOperationException($"Series '{series.Name}' must contain {expectedPointsPerSeries} data points, but found {series.DataPoints.Count}.");
            }
        }

        // Save the document.
        doc.Save("validated-chart.docx");
    }
}
