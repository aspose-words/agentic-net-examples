using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for the Shape class
using Aspose.Words.Drawing.Charts;        // Chart‑related APIs

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an empty column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Define categories (X‑axis labels) and series names (legend entries).
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        string[] seriesNames = { "Product A", "Product B", "Product C" };

        // Two‑dimensional array where rows correspond to series and columns to categories.
        double[,] data = new double[,]
        {
            { 120, 150, 170, 200 }, // Product A
            {  80, 130, 160, 190 }, // Product B
            { 100, 110, 130, 180 }  // Product C
        };

        int seriesCount = seriesNames.Length;
        int categoryCount = categories.Length;

        // Populate the chart with the custom data.
        for (int i = 0; i < seriesCount; i++)
        {
            double[] values = new double[categoryCount];
            for (int j = 0; j < categoryCount; j++)
            {
                values[j] = data[i, j];
            }

            // Add a series using the categories array and the values for this series.
            chart.Series.Add(seriesNames[i], categories, values);
        }

        // Optional: give the chart a title.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // Save the document.
        doc.Save("ChartFromArray.docx");
    }
}
