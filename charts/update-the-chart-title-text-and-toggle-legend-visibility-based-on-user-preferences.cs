using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // User preferences (hard‑coded for this example).
        string newTitle = "Quarterly Sales Report";
        bool showLegend = false; // Set to true to display the legend.

        // Create a new document and insert a sample chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        if (!chartShape.HasChart)
            throw new InvalidOperationException("Inserted shape does not contain a chart.");

        // Access the chart object.
        Chart chart = chartShape.Chart;

        // Update the chart title.
        ChartTitle title = chart.Title;
        title.Text = newTitle;
        title.Show = true; // Ensure the title is visible.

        // Toggle legend visibility based on the preference.
        ChartLegend legend = chart.Legend;
        legend.Position = showLegend ? LegendPosition.Right : LegendPosition.None;
        // Optionally allow other elements to overlap when the legend is hidden.
        legend.Overlay = !showLegend;

        // Save the document to the working directory.
        doc.Save("UpdatedChart.docx");
    }
}
