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

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the demo series that come with a newly inserted chart.
        chart.Series.Clear();

        // Define categories (X‑axis labels) and values for a new series.
        string[] categories = { "Category 1", "Category 2", "Category 3", "Category 4" };
        double[] values = { 10, 20, 30, 40 };

        // Add a new series with the categories and values.
        ChartSeries series = chart.Series.Add("Series 1", categories, values);

        // Define a distinct color for each data point.
        Color[] pointColors = { Color.Red, Color.Green, Color.Blue, Color.Orange };

        // Apply the colors to the individual data points.
        for (int i = 0; i < series.DataPoints.Count && i < pointColors.Length; i++)
        {
            ChartDataPoint dataPoint = series.DataPoints[i];
            dataPoint.Format.Fill.Color = pointColors[i];
        }

        // Save the document with the customized chart.
        doc.Save("AddDataPointsColors.docx");
    }
}
