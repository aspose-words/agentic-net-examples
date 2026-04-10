using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape shape = builder.InsertChart(ChartType.Column, 500, 300);

        // Ensure the shape actually contains a chart.
        if (!shape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart = shape.Chart;

        // Clear the demo data and add a simple custom series.
        chart.Series.Clear();
        chart.Series.Add("Sample Series",
            new[] { "Category 1", "Category 2", "Category 3" },
            new double[] { 10, 20, 30 });

        // Access the primary X‑axis.
        ChartAxis xAxis = chart.AxisX;

        // Show major gridlines.
        xAxis.HasMajorGridlines = true;

        // Customize the gridlines' line color and thickness.
        xAxis.Format.Stroke.Color = Color.DarkGray;
        xAxis.Format.Stroke.Weight = 1.5; // Thickness in points.

        // Save the document.
        doc.Save("ChartMajorGridlines.docx");
    }
}
