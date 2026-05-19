using System;
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

        // Remove the default demo data.
        chart.Series.Clear();

        // Add a simple series with categories and values.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120, 150, 180, 130 };
        chart.Series.Add("Sales", categories, values);

        // Set chart title and rotate it for visual emphasis.
        chart.Title.Text = "Quarterly Sales (3‑D Rotation)";
        chart.Title.Show = true;
        chart.Title.Rotation = 30; // Rotate title 30 degrees.

        // Rotate X‑axis tick labels.
        chart.AxisX.TickLabels.Orientation = ShapeTextOrientation.VerticalFarEast;
        chart.AxisX.TickLabels.Rotation = -30;

        // Rotate Y‑axis tick labels in the opposite direction.
        chart.AxisY.TickLabels.Orientation = ShapeTextOrientation.VerticalFarEast;
        chart.AxisY.TickLabels.Rotation = 30;

        // Save the document.
        doc.Save("3D_Rotation_ColumnChart.docx");
    }
}
