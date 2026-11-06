using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        if (!chartShape.HasChart)
        {
            throw new InvalidOperationException("Inserted shape does not contain a chart.");
        }

        // Access the chart.
        Chart chart = chartShape.Chart;

        // Clear any default series and add our own data.
        chart.Series.Clear();
        chart.Series.Add("Series 1", new double[] { 10, 20, 30 });
        chart.Series.Add("Series 2", new double[] { 15, 25, 35 });

        // Expected counts.
        const int expectedSeriesCount = 2;
        const int expectedPointsPerSeries = 3;

        // Validate series count.
        if (chart.Series.Count != expectedSeriesCount)
        {
            throw new InvalidOperationException(
                $"Chart must contain {expectedSeriesCount} series, but found {chart.Series.Count}.");
        }

        // Validate data point count for each series.
        foreach (ChartSeries series in chart.Series)
        {
            int pointsCount = series.DataPoints?.Count ?? 0;
            if (pointsCount != expectedPointsPerSeries)
            {
                throw new InvalidOperationException(
                    $"Series '{series.Name}' must contain {expectedPointsPerSeries} data points, but found {pointsCount}.");
            }
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "validated-chart.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
