using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ChartValidationExample
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Define categories that will be used for all series.
        string[] categories = { "Category 1", "Category 2", "Category 3" };

        // Add series with matching category/value counts.
        chart.Series.Add("Series A", categories, new double[] { 10, 20, 30 });
        chart.Series.Add("Series B", categories, new double[] { 15, 25, 35 });

        // Validate that every series has the same number of Y‑values as there are categories.
        foreach (ChartSeries series in chart.Series)
        {
            // For category‑based charts the numeric values are stored in YValues.
            int valuesCount = series.YValues.Count;
            if (valuesCount != categories.Length)
            {
                throw new InvalidOperationException(
                    $"Series '{series.Name}' contains {valuesCount} values, but the chart expects {categories.Length} categories.");
            }
        }

        // Save the document. The validation runs before this point.
        doc.Save("validated-chart.docx");
    }
}
