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

        // Insert a column chart with a defined size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the demo data that comes with a newly inserted chart.
        chart.Series.Clear();

        // Define categories (X‑axis labels) and series names (legend entries).
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        string[] seriesNames = { "Product A", "Product B", "Product C" };

        // Two‑dimensional array where rows represent series and columns represent category values.
        double[,] data = new double[,]
        {
            { 120.5, 135.0, 150.2, 165.3 }, // Product A
            { 80.0,  95.5, 110.0, 125.7 }, // Product B
            { 60.3,  70.1,  85.4,  95.0 }  // Product C
        };

        // Populate the chart with the custom data source.
        for (int seriesIndex = 0; seriesIndex < seriesNames.Length; seriesIndex++)
        {
            double[] values = new double[categories.Length];
            for (int catIndex = 0; catIndex < categories.Length; catIndex++)
            {
                values[catIndex] = data[seriesIndex, catIndex];
            }

            // Add a series using the categories array and the extracted values.
            chart.Series.Add(seriesNames[seriesIndex], categories, values);
        }

        // Save the document containing the chart.
        doc.Save("CustomDataChart.docx");
    }
}
