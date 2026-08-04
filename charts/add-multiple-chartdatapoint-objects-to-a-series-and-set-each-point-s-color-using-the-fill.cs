using System;
using System.Drawing;
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

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Define categories and values.
        string[] categories = { "A", "B", "C", "D", "E" };
        double[] values = { 10, 20, 30, 40, 50 };

        // Add a new series.
        ChartSeries series = chart.Series.Add("Series 1", categories, values);

        // Colors for each data point.
        Color[] pointColors = { Color.Red, Color.Green, Color.Blue, Color.Orange, Color.Purple };

        // Apply colors to individual data points.
        for (int i = 0; i < series.DataPoints.Count && i < pointColors.Length; i++)
        {
            ChartDataPoint point = series.DataPoints[i];
            point.Format.Fill.Color = pointColors[i];
        }

        // Save the document.
        doc.Save("ChartDataPointsColors.docx");
    }
}
