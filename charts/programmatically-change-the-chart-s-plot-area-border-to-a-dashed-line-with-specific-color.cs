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

        // Apply a dashed red border with a width of 2 points to the chart area.
        // Aspose.Words does not expose a PlotArea object directly, so we format the
        // overall chart area instead, which includes the plot area.
        chart.Format.Stroke.DashStyle = DashStyle.ShortDash;
        chart.Format.Stroke.Color = Color.Red;
        chart.Format.Stroke.Weight = 2.0;

        // Save the document to the working directory.
        doc.Save("ChartPlotAreaBorder.docx");
    }
}
