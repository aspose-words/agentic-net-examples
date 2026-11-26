using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Clear default demo series and add a custom series.
        chart.Series.Clear();
        chart.Series.Add("Series 1",
            new[] { "Category A", "Category B", "Category C" },
            new double[] { 10, 20, 30 });

        // Access the primary X‑axis.
        ChartAxis xAxis = chart.AxisX;

        // Show major gridlines.
        xAxis.HasMajorGridlines = true;

        // Customize gridline appearance: set color and thickness.
        xAxis.Format.Stroke.Color = Color.Blue;
        xAxis.Format.Stroke.Weight = 2.0;

        // Save the document.
        doc.Save("SetMajorGridlines.docx");
    }
}
