using System;
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

        // Remove the demo data that comes with a newly inserted chart.
        chart.Series.Clear();

        // Define categories that will be shared by all series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };

        // Define values for two series.
        double[] series1Values = { 10.0, 20.0, 30.0, 40.0 };
        double[] series2Values = { 15.0, 25.0, 35.0, 45.0 };

        // Validate that each series has the same number of values as categories.
        ValidateSeriesData(categories, series1Values);
        ValidateSeriesData(categories, series2Values);

        // Add the series to the chart.
        chart.Series.Add("Series 1", categories, series1Values);
        chart.Series.Add("Series 2", categories, series2Values);

        // Save the document.
        doc.Save("ValidatedChart.docx");
    }

    // Throws an exception if the number of values does not match the number of categories.
    private static void ValidateSeriesData(string[] categories, double[] values)
    {
        if (categories == null) throw new ArgumentNullException(nameof(categories));
        if (values == null) throw new ArgumentNullException(nameof(values));

        if (categories.Length != values.Length)
            throw new InvalidOperationException(
                $"Series data mismatch: categories count ({categories.Length}) does not match values count ({values.Length}).");
    }
}
