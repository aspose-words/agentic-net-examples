using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace ChartFrom2DArrayExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart with a defined size.
            Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
            Chart chart = chartShape.Chart;

            // Remove the demo data that Aspose.Words inserts by default.
            chart.Series.Clear();

            // Define categories (X‑axis labels).
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };

            // Define series names.
            string[] seriesNames = { "Product A", "Product B", "Product C" };

            // Two‑dimensional array holding the values.
            // Rows correspond to series, columns correspond to categories.
            double[,] data = new double[,]
            {
                { 120.5, 135.0, 150.2, 160.8 }, // Product A
                {  80.3,  95.4, 110.1, 115.6 }, // Product B
                {  60.0,  70.5,  85.3,  90.2 }  // Product C
            };

            // Add each series to the chart using the overload that accepts
            // a series name, an array of categories and an array of values.
            for (int i = 0; i < seriesNames.Length; i++)
            {
                double[] values = new double[categories.Length];
                for (int j = 0; j < categories.Length; j++)
                {
                    values[j] = data[i, j];
                }

                chart.Series.Add(seriesNames[i], categories, values);
            }

            // Save the document.
            doc.Save("ChartFrom2DArray.docx");
        }
    }
}
