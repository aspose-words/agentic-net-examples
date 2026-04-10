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

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Add a simple data series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120, 150, 180, 130 };
        chart.Series.Add("Sales", categories, values);

        // NOTE:
        // Aspose.Words does not expose direct properties for 3‑D rotation
        // (such as RotationX, RotationY, DepthPercent). These settings are
        // managed internally by the library for 3‑D chart types.
        // Therefore, no explicit rotation code is required here.

        // Save the document.
        doc.Save("3DRotationColumnChart.docx");
    }
}
