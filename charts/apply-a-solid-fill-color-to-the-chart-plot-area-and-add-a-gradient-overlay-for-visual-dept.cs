using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series and add custom data.
        chart.Series.Clear();
        string[] categories = new[] { "Category A", "Category B", "Category C" };
        chart.Series.Add("Series 1", categories, new double[] { 10, 20, 30 });

        // Apply a solid fill color to the chart area (which includes the plot area).
        chart.Format.Fill.Solid(Color.LightBlue);

        // Add a two‑color gradient overlay.
        // The foreground color is the solid fill set above.
        // The background color is a semi‑transparent white to create the overlay effect.
        chart.Format.Fill.BackColor = Color.FromArgb(128, Color.White);
        chart.Format.Fill.TwoColorGradient(GradientStyle.DiagonalDown, GradientVariant.Variant1);

        // Save the document.
        doc.Save("ChartPlotAreaFill.docx");
    }
}
