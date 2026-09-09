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

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that comes with a newly inserted chart.
        chart.Series.Clear();

        // Define a common set of categories for all series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };

        // Add series that correctly match the number of categories.
        chart.Series.Add("Series A", categories, new double[] { 10, 20, 30, 40 });
        chart.Series.Add("Series B", categories, new double[] { 15, 25, 35, 45 });

        // Uncomment the following line to see the validation exception in action
        // chart.Series.Add("Series Bad", categories, new double[] { 5, 10, 15 }); // Mismatched count

        // Validate that every series has the same number of data points as there are categories.
        ValidateSeriesCategoryAlignment(chart, categories.Length);

        // Save the resulting document.
        doc.Save("ValidatedChart.docx");
    }

    // Throws an exception if any series does not contain the expected number of values.
    private static void ValidateSeriesCategoryAlignment(Chart chart, int expectedCategoryCount)
    {
        foreach (ChartSeries series in chart.Series)
        {
            // For category‑based charts the YValues collection holds the data points.
            int valuesCount = series.YValues.Count;

            if (valuesCount != expectedCategoryCount)
            {
                throw new InvalidOperationException(
                    $"Series \"{series.Name}\" contains {valuesCount} values, " +
                    $"but expected {expectedCategoryCount} to match the category count.");
            }
        }
    }
}
