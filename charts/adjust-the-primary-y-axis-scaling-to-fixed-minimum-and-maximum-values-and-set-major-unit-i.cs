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
        Chart chart = chartShape.Chart;

        // Ensure we are working with a real chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Remove the demo data series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a custom data series.
        chart.Series.Add(
            "Sample Series",
            new[] { "Category A", "Category B", "Category C", "Category D" },
            new double[] { 10, 40, 70, 90 });

        // Adjust the primary Y‑axis scaling.
        ChartAxis yAxis = chart.AxisY;

        // Set fixed minimum and maximum values.
        yAxis.Scaling.Minimum = new AxisBound(0);   // Minimum value = 0
        yAxis.Scaling.Maximum = new AxisBound(100); // Maximum value = 100

        // Define the major unit interval (distance between major tick marks).
        yAxis.MajorUnit = 20; // Major ticks every 20 units

        // Optionally set a minor unit interval.
        yAxis.MinorUnit = 5; // Minor ticks every 5 units

        // Save the document to the local file system.
        doc.Save("ChartWithFixedYAxis.docx");
    }
}
