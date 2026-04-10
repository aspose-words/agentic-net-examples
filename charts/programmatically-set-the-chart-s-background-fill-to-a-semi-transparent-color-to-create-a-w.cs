using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a defined size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Optional: clear the default demo series and add custom data.
        chart.Series.Clear();
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Sales", categories, new double[] { 15000, 20000, 18000, 22000 });

        // Set the chart background fill to a semi‑transparent light gray.
        chart.Format.Fill.Solid(Color.LightGray);
        // Transparency value is between 0 (opaque) and 1 (fully transparent).
        chart.Format.Fill.Transparency = 0.5; // 50 % transparent – creates a watermark effect.

        // Save the document to the local file system.
        doc.Save("ChartWithWatermarkBackground.docx");
    }
}
