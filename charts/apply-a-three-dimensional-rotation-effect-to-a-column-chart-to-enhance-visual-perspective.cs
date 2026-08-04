using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

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

        // Add custom data for the chart.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Sales", categories, new double[] { 15000, 21000, 18000, 24000 });

        // Set a chart title.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // Show and rotate the X‑axis title to give a perspective effect.
        chart.AxisX.Title.Show = true;
        chart.AxisX.Title.Text = "Quarter";
        chart.AxisX.Title.Rotation = 45; // Rotate 45 degrees.

        // Show and rotate the Y‑axis title.
        chart.AxisY.Title.Show = true;
        chart.AxisY.Title.Text = "Revenue (USD)";
        chart.AxisY.Title.Rotation = -45; // Rotate -45 degrees.

        // Rotate the tick labels on both axes for additional 3‑D visual effect.
        chart.AxisX.TickLabels.Rotation = 30;
        chart.AxisY.TickLabels.Rotation = -30;

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "3D_Rotation_ColumnChart.docx");
        doc.Save(outputPath);
    }
}
