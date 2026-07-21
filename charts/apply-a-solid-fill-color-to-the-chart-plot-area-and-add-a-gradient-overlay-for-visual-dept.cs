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

        // Add custom data.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Revenue", categories, new double[] { 15000, 20000, 18000, 22000 });

        // Apply a solid fill to the chart area (used as a substitute for PlotArea).
        chart.Format.Fill.Solid(Color.LightGray);

        // Configure a two‑color vertical gradient overlay.
        // ForeColor is the start color, BackColor is the end color.
        chart.Format.Fill.ForeColor = Color.FromArgb(128, Color.LightBlue); // semi‑transparent light blue
        chart.Format.Fill.BackColor = Color.FromArgb(64, Color.DarkBlue);   // more transparent dark blue
        chart.Format.Fill.TwoColorGradient(GradientStyle.Vertical, GradientVariant.Variant1);

        // Save the document.
        doc.Save("ChartPlotAreaFill.docx");
    }
}
