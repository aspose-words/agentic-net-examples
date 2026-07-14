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

        // Change the chart's border (applies to the whole chart area) to a dashed red line with a width of 2 points.
        // Aspose.Words for .NET does not expose a separate PlotArea object in this version,
        // so we format the chart's outer border via the Chart.Format property.
        chart.Format.Stroke.Color = Color.Red;
        chart.Format.Stroke.Weight = 2.0;
        chart.Format.Stroke.DashStyle = DashStyle.Dash;

        // Save the document to the working directory.
        doc.Save("ChartPlotAreaBorder.docx");
    }
}
