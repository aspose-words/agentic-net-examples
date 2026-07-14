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

        // Insert a column chart and obtain its shape.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Clear the demo series and add a custom one.
        chart.Series.Clear();
        chart.Series.Add(
            "Sample Series",
            new[] { "Category 1", "Category 2", "Category 3" },
            new double[] { 10, 20, 30 });

        // Access the primary X axis.
        ChartAxis xAxis = chart.AxisX;

        // Make major gridlines visible.
        xAxis.HasMajorGridlines = true;

        // Set gridline color and thickness.
        xAxis.Format.Stroke.Color = Color.Blue;
        xAxis.Format.Stroke.Weight = 1.5; // thickness in points

        // Save the document.
        doc.Save("ChartGridlines.docx");
    }
}
