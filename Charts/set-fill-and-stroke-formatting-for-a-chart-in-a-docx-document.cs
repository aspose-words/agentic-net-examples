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

        // Remove the default series that Aspose.Words adds.
        chart.Series.Clear();

        // Add custom series data.
        string[] categories = { "Category 1", "Category 2" };
        chart.Series.Add("Series 1", categories, new double[] { 10, 20 });
        chart.Series.Add("Series 2", categories, new double[] { 30, 40 });

        // ----- Fill formatting -----
        // Set the chart background to a solid dark slate gray color.
        chart.Format.Fill.Solid(Color.DarkSlateGray);

        // ----- Stroke (line) formatting -----
        // Set the outline (stroke) color of the chart.
        chart.Format.Stroke.Color = Color.Black;
        // Set the thickness of the outline.
        chart.Format.Stroke.Weight = 2.0; // points
        // Ensure the stroke is visible.
        chart.Format.Stroke.On = true;

        // Save the document to disk.
        doc.Save("ChartFormatted.docx");
    }
}
