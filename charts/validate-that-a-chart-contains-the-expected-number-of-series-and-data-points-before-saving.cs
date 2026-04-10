using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ChartValidationExample
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words adds by default.
        chart.Series.Clear();

        // Define categories and data for two series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] series1Values = { 10.5, 12.3, 9.8, 14.2 };
        double[] series2Values = { 8.1, 11.0, 7.5, 13.4 };

        // Add the series.
        chart.Series.Add("Revenue 2023", categories, series1Values);
        chart.Series.Add("Revenue 2024", categories, series2Values);

        // Expected counts.
        const int expectedSeriesCount = 2;
        int expectedPointsPerSeries = categories.Length; // Not a compile‑time constant.

        // Validate series count.
        if (chart.Series.Count != expectedSeriesCount)
            throw new InvalidOperationException($"Expected {expectedSeriesCount} series, but found {chart.Series.Count}.");

        // Validate data point count for each series.
        foreach (ChartSeries series in chart.Series)
        {
            if (series.DataPoints.Count != expectedPointsPerSeries)
                throw new InvalidOperationException(
                    $"Series '{series.Name}' expected {expectedPointsPerSeries} data points, but found {series.DataPoints.Count}.");
        }

        // Save the document.
        doc.Save("ChartValidation.docx");
    }
}
