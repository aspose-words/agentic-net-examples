using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for the Shape class
using Aspose.Words.Drawing.Charts;        // Chart‑related APIs

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

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Define categories (X‑axis labels) and series (legend entries).
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        string[] seriesNames = { "Product A", "Product B", "Product C" };

        // Two‑dimensional array where rows correspond to series and columns to categories.
        double[,] values = {
            { 120.5, 135.0, 150.2, 165.3 }, // Product A
            { 80.0,  95.5,  110.1, 130.0 }, // Product B
            { 60.3,  70.4,  85.6,  95.2 }   // Product C
        };

        // Add each series to the chart using the categories and the corresponding row of values.
        for (int i = 0; i < seriesNames.Length; i++)
        {
            double[] seriesValues = new double[categories.Length];
            for (int j = 0; j < categories.Length; j++)
            {
                seriesValues[j] = values[i, j];
            }

            chart.Series.Add(seriesNames[i], categories, seriesValues);
        }

        // Optional: give the chart a title.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // Save the document.
        doc.Save("ChartFrom2DArray.docx");
    }
}
