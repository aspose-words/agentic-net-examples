using System;
using System.Drawing;
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

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Add a custom series with categories and values.
        chart.Series.Add("Sample Series",
            new[] { "A", "B", "C", "D" },
            new[] { 10.0, 20.0, 15.0, 25.0 });

        // Hide the plot area border by making its line invisible.
        chart.Format.Stroke.Color = Color.Transparent;
        chart.Format.Stroke.Weight = 0;

        // Axis lines remain visible (default behavior).

        // Save the document.
        doc.Save("PlotAreaBorderHidden.docx");
    }
}
