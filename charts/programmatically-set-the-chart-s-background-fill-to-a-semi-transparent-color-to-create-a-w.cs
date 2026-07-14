using System;
using System.Drawing;
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

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Optional: clear the demo series and add custom data.
        chart.Series.Clear();
        string[] categories = { "A", "B", "C" };
        chart.Series.Add("Series 1", categories, new double[] { 10, 20, 30 });

        // Set the chart background to a semi‑transparent light gray color.
        chart.Format.Fill.Solid(Color.LightGray);
        chart.Format.Fill.Transparency = 0.5; // 50 % transparency

        // Save the document.
        doc.Save("ChartWatermark.docx");
    }
}
