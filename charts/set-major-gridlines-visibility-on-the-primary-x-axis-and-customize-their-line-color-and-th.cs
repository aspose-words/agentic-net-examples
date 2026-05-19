using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Clear the demo data series.
        chart.Series.Clear();

        // Add a simple series so the chart has data.
        chart.Series.Add("Sample Series",
            new[] { "A", "B", "C", "D" },
            new double[] { 10, 20, 30, 40 });

        // Access the primary X axis.
        ChartAxis xAxis = chart.AxisX;

        // Show major gridlines.
        xAxis.HasMajorGridlines = true;

        // Customize gridline appearance: set line color and thickness.
        xAxis.Format.Stroke.Color = Color.DarkGray;
        xAxis.Format.Stroke.Weight = 1.5; // Thickness in points.

        // Save the document.
        doc.Save("SetMajorGridlines.docx");
    }
}
