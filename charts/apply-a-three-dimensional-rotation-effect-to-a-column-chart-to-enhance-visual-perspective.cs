using System;
using System.IO;
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

        // Insert a 3‑D column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column3D, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Add a simple data series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120, 150, 180, 130 };
        chart.Series.Add("Sales", categories, values);

        // NOTE:
        // Aspose.Words for .NET does not expose a View3D property for charts,
        // therefore a direct 3‑D rotation cannot be set via the API.
        // The chart type (Column3D) already provides a three‑dimensional appearance.

        // Optional: give the chart a title.
        chart.Title.Text = "Quarterly Sales (3‑D)";
        chart.Title.Show = true;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "3d-rotation-chart.docx");
        doc.Save(outputPath);
    }
}
