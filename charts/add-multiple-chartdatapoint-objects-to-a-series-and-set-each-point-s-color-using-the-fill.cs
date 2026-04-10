using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace ChartDataPointsExample
{
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

            // Remove the demo data that Aspose.Words inserts by default.
            chart.Series.Clear();

            // Define categories (X‑axis labels) and corresponding values.
            string[] categories = { "Category A", "Category B", "Category C", "Category D" };
            double[] values = { 10, 20, 30, 40 };

            // Add a single series with the categories and values.
            ChartSeries series = chart.Series.Add("Series 1", categories, values);

            // Define a distinct color for each data point.
            Color[] pointColors = { Color.Red, Color.Green, Color.Blue, Color.Orange };

            // Apply the colors to the individual data points.
            for (int i = 0; i < series.DataPoints.Count && i < pointColors.Length; i++)
            {
                series.DataPoints[i].Format.Fill.ForeColor = pointColors[i];
            }

            // Save the document containing the formatted chart.
            doc.Save("ChartDataPointsColors.docx");
        }
    }
}
