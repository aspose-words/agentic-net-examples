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
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Verify that the shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Access the primary X‑axis.
        ChartAxis xAxis = chart.AxisX;

        // Show major gridlines.
        xAxis.HasMajorGridlines = true;

        // Customize gridline appearance: color and thickness.
        xAxis.Format.Stroke.Color = Color.DarkGray;
        xAxis.Format.Stroke.Weight = 0.75; // Thickness in points.

        // Save the document.
        doc.Save("ChartMajorGridlines.docx");
    }
}
