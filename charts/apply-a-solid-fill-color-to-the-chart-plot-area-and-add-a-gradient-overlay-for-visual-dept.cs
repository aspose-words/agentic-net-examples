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
        if (!chartShape.HasChart)
            throw new InvalidOperationException("Inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Remove the demo data series and add custom data.
        chart.Series.Clear();
        string[] categories = { "A", "B", "C" };
        chart.Series.Add("Series 1", categories, new double[] { 10, 20, 30 });

        // Apply a solid fill color to the chart area (plot area is not directly exposed in Aspose.Words).
        chart.Format.Fill.Solid(Color.LightBlue);

        // Add a vertical two‑color gradient overlay for visual depth.
        chart.Format.Fill.TwoColorGradient(GradientStyle.Vertical, GradientVariant.Variant1);

        // Save the document.
        doc.Save("ChartPlotAreaFill.docx");
    }
}
