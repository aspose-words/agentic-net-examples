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

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Aspose.Words does not expose a separate PlotArea object.
        // The border of the chart's plot area can be styled through the
        // chart's overall format (Chart.Format). This will affect the
        // visible border around the plotted data region.
        chart.Format.Stroke.Color = Color.Red;          // Red border color
        chart.Format.Stroke.DashStyle = DashStyle.Dash; // Dashed line style
        chart.Format.Stroke.Weight = 2.0;               // Width of 2 points

        // Save the document.
        doc.Save("ChartPlotAreaBorder.docx");
    }
}
