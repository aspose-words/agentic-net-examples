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
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the demo data.
        chart.Series.Clear();

        // Define categories that will be used for all series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        int expectedCount = categories.Length;

        // Add two series with matching value counts.
        chart.Series.Add("Sales 2022", categories, new double[] { 120, 150, 130, 170 });
        chart.Series.Add("Sales 2023", categories, new double[] { 140, 160, 150, 180 });

        // Validation: ensure every series has the same number of Y values as the categories.
        foreach (ChartSeries series in chart.Series)
        {
            // YValues holds the numeric data for the series.
            int valueCount = series.YValues.Count;

            if (valueCount != expectedCount)
                throw new InvalidOperationException(
                    $"Series '{series.Name}' contains {valueCount} values, but {expectedCount} categories are required.");
        }

        // Save the document – if validation passes, the file is written.
        doc.Save("validated-chart.docx");
    }
}
