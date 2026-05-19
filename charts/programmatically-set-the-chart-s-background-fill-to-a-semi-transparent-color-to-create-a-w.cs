using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add sample data to the chart.
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        chart.Series.Add("Sample Series", categories, new double[] { 15, 30, 45 });

        // Apply a semi‑transparent background fill to create a watermark effect.
        chart.Format.Fill.Solid(Color.LightGray);
        chart.Format.Fill.Transparency = 0.5; // 0 = fully opaque, 1 = fully transparent

        // Save the document.
        doc.Save("ChartWithWatermarkBackground.docx");
    }
}
