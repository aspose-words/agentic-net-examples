using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ChartSeriesValidationExample
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Ensure the chart starts with no series.
        chart.Series.Clear();

        // Add two series with mismatched category counts.
        // Series 1 has three categories.
        chart.Series.Add("Series 1",
            new[] { "Category A", "Category B", "Category C" },
            new[] { 10.0, 20.0, 30.0 });

        // Series 2 has only two categories (mismatch).
        chart.Series.Add("Series 2",
            new[] { "Category A", "Category B" },
            new[] { 15.0, 25.0 });

        // Validate the chart series.
        try
        {
            ValidateSeriesCategoryCounts(chart);
            Console.WriteLine("Validation passed: all series have matching category counts.");
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine($"Validation failed: {ex.Message}");
            Console.WriteLine("Fixing the mismatch by rebuilding the series with consistent categories.");

            // Rebuild series with consistent categories.
            string[] categories = new[] { "Category A", "Category B", "Category C" };
            chart.Series.Clear();

            chart.Series.Add("Series 1", categories, new[] { 10.0, 20.0, 30.0 });
            chart.Series.Add("Series 2", categories, new[] { 15.0, 25.0, 0.0 }); // Added missing value.

            // Re‑validate after fixing.
            ValidateSeriesCategoryCounts(chart);
            Console.WriteLine("After fixing, validation passed.");
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ValidatedChart.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    private static void ValidateSeriesCategoryCounts(Chart chart)
    {
        if (chart == null) throw new ArgumentNullException(nameof(chart));
        if (chart.Series == null || chart.Series.Count == 0)
            throw new InvalidOperationException("Chart must contain at least one series.");

        // Expected number of categories is taken from the first series.
        int expectedCount = chart.Series[0].DataPoints?.Count ?? 0;

        foreach (ChartSeries series in chart.Series)
        {
            int actualCount = series.DataPoints?.Count ?? 0;
            if (actualCount != expectedCount)
            {
                throw new InvalidOperationException(
                    $"Series \"{series.Name}\" has {actualCount} data points, " +
                    $"but expected {expectedCount} to match other series.");
            }
        }
    }
}
