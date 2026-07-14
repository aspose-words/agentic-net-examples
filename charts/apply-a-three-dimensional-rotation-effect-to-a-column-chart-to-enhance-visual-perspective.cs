using System;
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

        // Remove the default demo data.
        chart.Series.Clear();

        // Add a simple series with categories and values.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120, 150, 180, 130 };
        chart.Series.Add("Sales", categories, values);

        // NOTE:
        // Aspose.Words does not expose direct RotationX/RotationY properties for 3‑D charts.
        // The chart is already a 3‑D column chart, which provides a default perspective.
        // If future versions add rotation support, the appropriate properties can be set here.

        // Save the document.
        doc.Save("3D_Rotated_Column_Chart.docx");
    }
}
