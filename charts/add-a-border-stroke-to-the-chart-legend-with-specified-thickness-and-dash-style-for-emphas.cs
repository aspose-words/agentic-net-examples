using System;
using System.IO;
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

        // Insert a column chart with a defined size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Clear the default demo series.
        chart.Series.Clear();

        // Add sample data to the chart.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Sales 2023", categories, new double[] { 15000, 20000, 18000, 22000 });
        chart.Series.Add("Sales 2024", categories, new double[] { 17000, 21000, 19000, 23000 });

        // Configure the legend border (stroke) for emphasis.
        // Set the border thickness.
        chart.Legend.Format.Stroke.Weight = 2.0; // Thickness in points.

        // Set the dash style of the border.
        chart.Legend.Format.Stroke.DashStyle = DashStyle.Dash; // Dashed line.

        // Optionally set the border color.
        chart.Legend.Format.Stroke.Color = Color.DarkRed;

        // Ensure the legend is visible and positioned.
        chart.Legend.Position = LegendPosition.Right;
        chart.Legend.Overlay = false;

        // Save the document to the local file system.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "ChartLegendBorder.docx");
        doc.Save(outputPath);
    }
}
