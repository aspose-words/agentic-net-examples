using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

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

        // Clear default demo data and add custom series.
        chart.Series.Clear();
        chart.Series.Add("Sample Series",
            new[] { "A", "B", "C", "D" },
            new double[] { 10, 20, 30, 40 });

        // Access the primary X axis.
        ChartAxis xAxis = chart.AxisX;

        // Make major gridlines visible.
        xAxis.HasMajorGridlines = true;

        // Customize gridline appearance: set line color and thickness.
        xAxis.Format.Stroke.Color = Color.Blue;
        xAxis.Format.Stroke.Weight = 2.0; // Thickness in points.

        // Save the document.
        doc.Save("ChartWithCustomGridlines.docx");
    }
}
