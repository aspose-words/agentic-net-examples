using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // Ensure the shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Get the chart object.
        Chart chart = chartShape.Chart;

        // Change the plot area (chart area) border to a dashed red line with a width of 2 points.
        // The Chart.Format property provides access to the chart's line formatting.
        chart.Format.Stroke.DashStyle = DashStyle.Dash;
        chart.Format.Stroke.Color = Color.Red;
        chart.Format.Stroke.Weight = 2.0;

        // Save the document.
        doc.Save("PlotAreaBorder.docx");
    }
}
