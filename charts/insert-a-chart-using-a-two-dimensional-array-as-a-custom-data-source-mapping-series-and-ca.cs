using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for Shape
using Aspose.Words.Drawing.Charts;        // Chart APIs

namespace ChartFromArrayExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart.
            Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = chartShape.Chart;

            // Remove the default demo series.
            chart.Series.Clear();

            // X‑axis categories.
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };

            // Series (row) names.
            string[] seriesNames = { "Product A", "Product B", "Product C" };

            // Two‑dimensional data: rows = series, columns = categories.
            double[,] data = new double[,]
            {
                { 120.5, 135.0, 150.2, 160.8 }, // Product A
                {  95.3, 110.4, 115.6, 130.0 }, // Product B
                {  80.0,  85.5,  90.1, 100.2 }  // Product C
            };

            // Populate the chart with the data from the array.
            for (int i = 0; i < seriesNames.Length; i++)
            {
                double[] values = new double[categories.Length];
                for (int j = 0; j < categories.Length; j++)
                {
                    values[j] = data[i, j];
                }

                // Add a series: name, categories, values.
                chart.Series.Add(seriesNames[i], categories, values);
            }

            // Save the document containing the chart.
            doc.Save("ChartFromArray.docx");
        }
    }
}
