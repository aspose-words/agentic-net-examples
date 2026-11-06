using System;
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
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);

        // Ensure the inserted shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Access the chart object.
        Chart chart = chartShape.Chart;

        // Update the chart title.
        chart.Title.Text = "Quarterly Sales Overview";

        // User preference for legend visibility (hard‑coded for this example).
        bool showLegend = false;

        // Toggle legend visibility if the legend object exists.
        if (chart.Legend != null)
        {
            // Aspose.Words does not expose a direct Visible property for the legend.
            // As a workaround, we change the legend position and overlay setting.
            // When showLegend is false we move the legend to a position that effectively hides it.
            if (showLegend)
            {
                chart.Legend.Position = LegendPosition.Right;
                chart.Legend.Overlay = false;
            }
            else
            {
                // Position the legend at the bottom and overlay it so it does not appear prominently.
                chart.Legend.Position = LegendPosition.Bottom;
                chart.Legend.Overlay = true;
            }
        }

        // Save the updated document.
        doc.Save("updated-chart.docx");
    }
}
