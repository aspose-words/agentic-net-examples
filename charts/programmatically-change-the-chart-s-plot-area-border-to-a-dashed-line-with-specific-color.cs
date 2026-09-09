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

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // NOTE:
        // In Aspose.Words v2 the Chart class does not expose a PlotArea property.
        // Therefore we apply the border formatting to the chart area itself,
        // which is the closest available option.
        chart.Format.Stroke.Color = Color.Red;               // Set border color.
        chart.Format.Stroke.DashStyle = Aspose.Words.Drawing.DashStyle.Dash; // Dashed line.
        chart.Format.Stroke.Weight = 2.0;                    // Width in points.

        // Save the document.
        doc.Save("PlotAreaBorder.docx");
    }
}
