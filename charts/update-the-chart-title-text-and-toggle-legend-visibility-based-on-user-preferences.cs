using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class UpdateChartExample
{
    public static void Main()
    {
        // User preferences (hard‑coded for this example).
        bool showTitle = true;
        string newTitleText = "Quarterly Sales";
        bool showLegend = false; // Set to true to display the legend.

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Update the chart title according to preferences.
        ChartTitle title = chart.Title;
        title.Text = newTitleText;
        title.Show = showTitle;
        // Optional: allow other elements to overlap the title.
        title.Overlay = true;

        // Toggle legend visibility based on preferences.
        ChartLegend legend = chart.Legend;
        legend.Position = showLegend ? LegendPosition.Right : LegendPosition.None;
        // Optional: allow other elements to overlap the legend.
        legend.Overlay = true;

        // Save the document.
        doc.Save("UpdatedChart.docx");
    }
}
