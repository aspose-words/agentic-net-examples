using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Define categories (X‑axis) and values for two series.
        string[] categories = { "Q1", "Q2", "Q3" };
        double[] series1Values = { 10.5, 20.0, 15.2 };
        double[] series2Values = { 12.3, 18.7, 22.1 };

        // Add the two series to the chart.
        chart.Series.Add("Series 1", categories, series1Values);
        chart.Series.Add("Series 2", categories, series2Values);

        // Expected counts for validation.
        const int expectedSeriesCount = 2;
        const int expectedDataPointsPerSeries = 3;

        // Validate the number of series.
        if (chart.Series.Count != expectedSeriesCount)
            throw new InvalidOperationException($"Chart must contain {expectedSeriesCount} series, but found {chart.Series.Count}.");

        // Validate the number of data points in each series.
        foreach (ChartSeries series in chart.Series)
        {
            if (series.DataPoints.Count != expectedDataPointsPerSeries)
                throw new InvalidOperationException($"Series '{series.Name}' must contain {expectedDataPointsPerSeries} data points, but found {series.DataPoints.Count}.");
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "ValidatedChart.docx");
        doc.Save(outputPath);
    }
}
