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

        // Remove the demo data series.
        chart.Series.Clear();

        // Add a simple series so the chart has visible data.
        string[] categories = new[] { "Category 1", "Category 2", "Category 3" };
        chart.Series.Add("Series 1", categories, new double[] { 10, 20, 30 });

        // Apply a solid fill color to the chart area (plot area is not directly exposed in Aspose.Words).
        chart.Format.Fill.Solid(Color.LightBlue);

        // Add a gradient overlay for visual depth.
        chart.Format.Fill.TwoColorGradient(GradientStyle.DiagonalDown, GradientVariant.Variant2);
        chart.Format.Fill.BackColor = Color.White; // gradient from white to LightBlue

        // Save the document.
        doc.Save("ChartPlotAreaFill.docx");
    }
}
