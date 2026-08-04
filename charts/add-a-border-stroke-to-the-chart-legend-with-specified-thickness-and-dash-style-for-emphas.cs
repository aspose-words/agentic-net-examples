using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Clear the default demo series.
        chart.Series.Clear();

        // Add sample data to the chart.
        string[] categories = new[] { "Category 1", "Category 2", "Category 3" };
        chart.Series.Add("Series 1", categories, new double[] { 10, 20, 30 });
        chart.Series.Add("Series 2", categories, new double[] { 15, 25, 35 });

        // Access the chart legend.
        ChartLegend legend = chart.Legend;

        // Apply a border (stroke) to the legend.
        // Set the thickness (weight) of the border.
        legend.Format.Stroke.Weight = 2.0; // Thickness in points.

        // Set the dash style of the border.
        legend.Format.Stroke.DashStyle = DashStyle.Dash; // Dashed line.

        // Optionally set the border color.
        legend.Format.Stroke.Color = Color.DarkBlue;

        // Save the document to the working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartLegendBorder.docx");
        doc.Save(outputPath);
    }
}
