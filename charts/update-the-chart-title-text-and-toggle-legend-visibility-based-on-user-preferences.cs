using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // User preferences (could be loaded from a config file or passed as arguments)
        string preferredTitle = "Quarterly Sales Overview";
        bool showLegend = false; // Set to true to display the legend, false to hide it

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        // Ensure the shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Access the chart object.
        Chart chart = chartShape.Chart;

        // Update the chart title according to the user preference.
        ChartTitle title = chart.Title;
        title.Text = preferredTitle;
        title.Show = true; // Ensure the title is visible.

        // Toggle legend visibility based on the user preference.
        ChartLegend legend = chart.Legend;
        legend.Position = showLegend ? LegendPosition.Right : LegendPosition.None;

        // Optionally, adjust legend overlay to give more space to the chart.
        legend.Overlay = true;

        // Save the document.
        doc.Save("UpdatedChart.docx");
    }
}
