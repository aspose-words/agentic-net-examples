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
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add sample data to the chart.
        string[] categories = { "Category A", "Category B", "Category C" };
        chart.Series.Add("Sample Series", categories, new double[] { 15, 30, 45 });

        // Apply a solid fill color to the chart area (used as a fallback for plot area).
        chart.Format.Fill.Solid(Color.LightBlue);

        // Add a vertical two‑color gradient overlay for visual depth.
        chart.Format.Fill.ForeColor = Color.LightBlue;
        chart.Format.Fill.BackColor = Color.DarkBlue;
        chart.Format.Fill.TwoColorGradient(GradientStyle.Vertical, GradientVariant.Variant1);

        // Save the document.
        doc.Save("ChartPlotAreaFill.docx");
    }
}
