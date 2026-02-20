using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class ChartFormattingExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a specific size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Example: clear the default demo series and add custom data.
        chart.Series.Clear();
        chart.Series.Add("Series 1", new[] { "A", "B", "C" }, new double[] { 10, 20, 30 });

        // ----- Fill formatting -----
        // Set a solid fill color for the entire chart area.
        chart.Format.Fill.Solid(Color.LightBlue);

        // ----- Stroke (line) formatting -----
        // Set the outline (stroke) color, weight and dash style.
        chart.Format.Stroke.Color = Color.DarkBlue;   // Outline color.
        chart.Format.Stroke.Weight = 2.0;            // Thickness in points.
        chart.Format.Stroke.DashStyle = DashStyle.Dash; // Dashed line.

        // Save the document to a DOCX file.
        string artifactsDir = @"C:\Temp\";
        doc.Save(System.IO.Path.Combine(artifactsDir, "ChartWithFillAndStroke.docx"));
    }
}
