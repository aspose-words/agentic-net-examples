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

        // Insert a chart (Column type) with a reasonable size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Apply a semi‑transparent fill to the chart area to achieve a watermark effect.
        // First set a solid fill color, then adjust its transparency (0 = opaque, 1 = fully transparent).
        chart.Format.Fill.Solid(Color.LightGray);
        chart.Format.Fill.Transparency = 0.5; // 50 % transparent

        // Save the document.
        doc.Save("ChartBackgroundWatermark.docx");
    }
}
