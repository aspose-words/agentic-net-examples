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

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a custom data series.
        chart.Series.Add(
            "Sample Series",
            new[] { "Category 1", "Category 2", "Category 3" },
            new double[] { 15, 30, 45 });

        // Configure the primary Y‑axis: set fixed minimum, maximum and major unit.
        ChartAxis yAxis = chart.AxisY;
        yAxis.Scaling.Minimum = new AxisBound(0);   // Fixed minimum value.
        yAxis.Scaling.Maximum = new AxisBound(50);  // Fixed maximum value.
        yAxis.MajorUnit = 10;                       // Major tick interval.

        // Save the document with the modified chart.
        doc.Save("YaxisScalingChart.docx");
    }
}
