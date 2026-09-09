using System;
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

        // Insert a column chart with a reasonable size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series and add our own data.
        chart.Series.Clear();
        chart.Series.Add(
            "Sample Series",
            new[] { "A", "B", "C", "D", "E" },
            new double[] { 30, 120, 80, 150, 60 });

        // Adjust the primary Y‑axis scaling.
        ChartAxis yAxis = chart.AxisY;
        yAxis.Scaling.Minimum = new AxisBound(0);      // Fixed minimum value.
        yAxis.Scaling.Maximum = new AxisBound(200);    // Fixed maximum value.
        yAxis.MajorUnit = 50;                          // Major tick interval.

        // Optionally set minor unit for finer grid lines.
        yAxis.MinorUnit = 10;

        // Save the document.
        doc.Save("YaxisScaling.docx");
    }
}
