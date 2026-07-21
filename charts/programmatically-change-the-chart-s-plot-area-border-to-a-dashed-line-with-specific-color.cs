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
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Aspose.Words does not expose a dedicated PlotArea object.
        // The closest available formatting is the chart's overall format.
        // Here we set a dashed red border with a width of 2 points for the chart area.
        chart.Format.Stroke.DashStyle = DashStyle.Dash;
        chart.Format.Stroke.Color = Color.Red;
        chart.Format.Stroke.Weight = 2.0; // Width in points.

        // Save the document.
        doc.Save("PlotAreaBorder.docx");
    }
}
