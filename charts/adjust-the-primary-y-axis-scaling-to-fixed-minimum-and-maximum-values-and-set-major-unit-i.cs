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

        // Insert a column chart with a defined size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);

        // Verify that the shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Access the chart object.
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Add a simple series so the chart has visible data.
        chart.Series.Add("Sample Series",
            new[] { "A", "B", "C", "D", "E" },
            new double[] { 10, 30, 55, 70, 90 });

        // Adjust the primary Y‑axis scaling.
        ChartAxis yAxis = chart.AxisY;

        // Set fixed minimum and maximum values.
        yAxis.Scaling.Minimum = new AxisBound(0);    // Minimum = 0
        yAxis.Scaling.Maximum = new AxisBound(100); // Maximum = 100

        // Set major and minor unit intervals.
        yAxis.MajorUnit = 20; // Major tick every 20 units
        yAxis.MinorUnit = 5;  // Minor tick every 5 units

        // Save the document.
        doc.Save("YaxisScaling.docx");
    }
}
