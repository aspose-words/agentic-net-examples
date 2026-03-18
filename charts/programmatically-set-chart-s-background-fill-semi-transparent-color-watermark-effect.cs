using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

class ChartWatermarkExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart (any type, size is arbitrary).
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Optional: clear default series and add custom data.
        chart.Series.Clear();
        string[] categories = new string[] { "A", "B", "C" };
        chart.Series.Add("Series 1", categories, new double[] { 10, 20, 30 });

        // Set the chart background to a solid color.
        chart.Format.Fill.Solid(Color.LightGray);

        // Make the background semi‑transparent (50% transparent).
        // Transparency value is between 0.0 (opaque) and 1.0 (clear).
        chart.Format.Fill.Transparency = 0.5;

        // Save the document.
        doc.Save("ChartWithWatermark.docx");
    }
}
