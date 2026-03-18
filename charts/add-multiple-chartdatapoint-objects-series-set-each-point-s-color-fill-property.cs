using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing; // <-- added
using Aspose.Words.Drawing.Charts;

class ChartDataPointColorsExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 350);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Define categories (X axis) and corresponding Y values.
        string[] categories = new[] { "Category A", "Category B", "Category C", "Category D" };
        double[] values = new[] { 10.0, 20.0, 30.0, 40.0 };

        // Add a new series with the defined categories and values.
        ChartSeries series = chart.Series.Add("Series 1", categories, values);

        // Define a color for each data point.
        Color[] pointColors = new[] { Color.Red, Color.Green, Color.Blue, Color.Orange };

        // Apply the colors to the individual data points.
        for (int i = 0; i < series.DataPoints.Count; i++)
        {
            series.DataPoints[i].Format.Fill.Color = pointColors[i];
        }

        // Save the document containing the chart.
        doc.Save("ChartDataPointsColors.docx");
    }
}
