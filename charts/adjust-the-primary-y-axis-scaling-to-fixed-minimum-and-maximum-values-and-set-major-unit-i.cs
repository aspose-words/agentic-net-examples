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
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the default demo series to start with a clean chart.
        chart.Series.Clear();

        // Add a simple series with categories and values.
        chart.Series.Add(
            "Sample Series",
            new[] { "Category A", "Category B", "Category C", "Category D" },
            new double[] { 10, 30, 50, 70 });

        // Access the primary Y‑axis.
        ChartAxis yAxis = chart.AxisY;

        // Set fixed minimum and maximum bounds.
        yAxis.Scaling.Minimum = new AxisBound(0);    // Minimum value = 0
        yAxis.Scaling.Maximum = new AxisBound(100); // Maximum value = 100

        // Define the major unit interval (distance between major tick marks).
        yAxis.MajorUnit = 20; // Major tick every 20 units

        // Save the document with the modified chart.
        doc.Save("YaxisScaling.docx");
    }
}
