using System;
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
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);

        // Ensure the shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Define categories that will be shared by all series.
        string[] categories = { "Category 1", "Category 2", "Category 3" };

        // Add two correctly aligned series.
        chart.Series.Add("Series A", categories, new double[] { 10, 20, 30 });
        chart.Series.Add("Series B", categories, new double[] { 15, 25, 35 });

        // Validate that all series have the same number of categories.
        ValidateSeriesCategoryCounts(chart);

        // Add a series with correct data first.
        chart.Series.Add("Series C (Invalid)", categories, new double[] { 5, 15, 25 });

        // Introduce a mismatch by adding an extra X value without a corresponding Y value.
        ChartSeries invalidSeries = chart.Series[chart.Series.Count - 1];
        invalidSeries.Add(ChartXValue.FromString("Extra Category"));

        try
        {
            // This call will now throw because the last series has a different category count.
            ValidateSeriesCategoryCounts(chart);
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine($"Validation error: {ex.Message}");
        }

        // Save the document.
        doc.Save("ChartValidation.docx");
    }

    private static void ValidateSeriesCategoryCounts(Chart chart)
    {
        if (chart.Series.Count == 0)
            return; // No series to validate.

        // Use the first series as the reference for the expected category count.
        int expectedCount = chart.Series[0].XValues.Count;

        for (int i = 0; i < chart.Series.Count; i++)
        {
            ChartSeries series = chart.Series[i];
            int actualCount = series.XValues.Count;

            if (actualCount != expectedCount)
                throw new InvalidOperationException(
                    $"Series '{series.Name}' has {actualCount} categories, expected {expectedCount}.");
        }
    }
}
