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

        // Apply a semi‑transparent fill to the chart background to achieve a watermark effect.
        chart.Format.Fill.Solid(Color.LightGray);
        chart.Format.Fill.Transparency = 0.5; // 50 % transparency

        // Save the document.
        doc.Save("ChartBackgroundWatermark.docx");
    }
}
