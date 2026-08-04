using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace ChartPlotAreaFillExample
{
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

            // Remove the demo data.
            chart.Series.Clear();

            // Add a simple series so the chart has visible data.
            string[] categories = { "A", "B", "C" };
            chart.Series.Add("Series 1", categories, new double[] { 10, 20, 30 });

            // Apply a solid fill color to the chart area (used as plot area).
            chart.Format.Fill.Solid(Color.LightBlue);

            // Add a gradient overlay for visual depth.
            chart.Format.Fill.ForeColor = Color.LightBlue;
            chart.Format.Fill.BackColor = Color.DarkBlue;
            // Apply a vertical two‑color gradient.
            chart.Format.Fill.TwoColorGradient(GradientStyle.Vertical, GradientVariant.Variant1);

            // Save the document.
            doc.Save("ChartPlotAreaFill.docx");
        }
    }
}
