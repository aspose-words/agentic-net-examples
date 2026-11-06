using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove any default series.
        chart.Series.Clear();

        // Define values and corresponding colors for each data point.
        double[] values = { 10, 20, 30 };
        Color[] colors = { Color.Red, Color.Green, Color.Blue };

        // Add a series with the defined values.
        ChartSeries series = chart.Series.Add("Sample Series", values);

        // Set the fill color for each data point.
        for (int i = 0; i < series.DataPoints.Count && i < colors.Length; i++)
        {
            series.DataPoints[i].Format.Fill.ForeColor = colors[i];
        }

        // Save the document.
        doc.Save("ChartDataPoints.docx");
    }
}
