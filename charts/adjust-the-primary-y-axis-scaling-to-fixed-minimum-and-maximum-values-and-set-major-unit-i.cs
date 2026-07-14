using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;          // Needed for the Shape class
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

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a custom series with categories and values.
        chart.Series.Add(
            "Sample Series",
            new[] { "A", "B", "C", "D" },
            new double[] { 10, 30, 20, 40 });

        // Configure the primary Y‑axis scaling.
        ChartAxis yAxis = chart.AxisY;
        yAxis.Scaling.Minimum = new AxisBound(0);   // Fixed minimum value.
        yAxis.Scaling.Maximum = new AxisBound(50);  // Fixed maximum value.
        yAxis.MajorUnit = 10;                       // Major tick interval.

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "YaxisScaling.docx");
        doc.Save(outputPath);
    }
}
