using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Define categories and corresponding values.
        string[] categories = new[] { "Category 1", "Category 2", "Category 3" };
        double[] values = new[] { 10.0, 20.0, 30.0 };

        // Add a new series with the defined categories and values.
        ChartSeries series = chart.Series.Add("Series 1", categories, values);

        // Set a distinct fill color for each data point in the series.
        series.DataPoints[0].Format.Fill.Color = Color.Red;
        series.DataPoints[1].Format.Fill.Color = Color.Green;
        series.DataPoints[2].Format.Fill.Color = Color.Blue;

        // Save the document containing the customized chart.
        doc.Save("ChartDataPointsColors.docx");
    }
}
