using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;               // Required for Shape
using Aspose.Words.Drawing.Charts;        // Chart related types

public class ChartValidationExample
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Define categories that will be used for all series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };

        // Define values for two series. Ensure each values array matches the categories length.
        double[] salesValues = { 15000, 20000, 18000, 22000 };
        double[] profitValues = { 3000, 4000, 3500, 5000 };

        // Validation: category count must equal values count for each series.
        ValidateSeriesData(categories, salesValues);
        ValidateSeriesData(categories, profitValues);

        // Add the series to the chart.
        chart.Series.Add("Sales", categories, salesValues);
        chart.Series.Add("Profit", categories, profitValues);

        // After adding, verify that every series has the same number of data points as categories.
        foreach (ChartSeries series in chart.Series)
        {
            // YValues holds the numeric data for the series.
            // Its Count should match the number of categories.
            if (series.YValues.Count != categories.Length)
                throw new InvalidOperationException(
                    $"Series \"{series.Name}\" has {series.YValues.Count} values, expected {categories.Length}.");
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "validated-chart.docx");
        doc.Save(outputPath);
    }

    // Helper method to validate that a values array matches the categories array length.
    private static void ValidateSeriesData(string[] categories, double[] values)
    {
        if (categories == null) throw new ArgumentNullException(nameof(categories));
        if (values == null) throw new ArgumentNullException(nameof(values));

        if (categories.Length != values.Length)
            throw new InvalidOperationException(
                $"Category count ({categories.Length}) does not match values count ({values.Length}).");
    }
}
