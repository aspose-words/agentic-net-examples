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
        Shape chartShape = builder.InsertChart(ChartType.Column3D, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add sample data.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Sales", categories, new double[] { 150, 200, 180, 220 });

        // Note: Aspose.Words does not expose direct RotationX/RotationY properties for charts.
        // The 3‑D effect is already applied by using the Column3D chart type.

        // Save the document.
        doc.Save("3DColumnChart.docx");
    }
}
