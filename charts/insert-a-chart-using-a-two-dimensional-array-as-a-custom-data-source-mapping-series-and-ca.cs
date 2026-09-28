using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace ChartFromArrayExample
{
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

            // Remove any default series that Aspose.Words adds automatically.
            chart.Series.Clear();

            // Define series names, category names, and the numeric data matrix.
            string[] seriesNames = { "North", "South" };
            string[] categories = { "Q1", "Q2", "Q3" };
            double[,] values = {
                { 120.5, 135.0, 150.2 }, // North series values
                { 98.3, 110.7, 130.4 }   // South series values
            };

            // Populate the chart with series based on the two‑dimensional array.
            // Use the overload that also accepts category names, so explicit axis handling is unnecessary.
            for (int i = 0; i < seriesNames.Length; i++)
            {
                double[] seriesValues = new double[categories.Length];
                for (int j = 0; j < categories.Length; j++)
                {
                    seriesValues[j] = values[i, j];
                }

                // Add the series together with the category labels.
                chart.Series.Add(seriesNames[i], categories, seriesValues);
            }

            // Save the document containing the chart.
            doc.Save("ChartFromArray.docx");
        }
    }
}
