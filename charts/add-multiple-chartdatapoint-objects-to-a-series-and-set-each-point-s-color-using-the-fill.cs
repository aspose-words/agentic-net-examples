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

        // Remove the demo series that come with a new chart.
        chart.Series.Clear();

        // Define categories and values for a new series.
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        double[] values = { 10, 20, 30 };

        // Add the series to the chart.
        ChartSeries series = chart.Series.Add("Series 1", categories, values);

        // Define a set of colors to apply to the data points.
        Color[] pointColors = { Color.Red, Color.Green, Color.Blue };

        // Apply a different fill color to each data point.
        for (int i = 0; i < series.DataPoints.Count; i++)
        {
            ChartDataPoint point = series.DataPoints[i];
            point.Format.Fill.Color = pointColors[i % pointColors.Length];
        }

        // Save the document with the customized chart.
        doc.Save("ChartDataPointsColors.docx");
    }
}
