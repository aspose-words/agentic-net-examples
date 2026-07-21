using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // User preferences (could be loaded from a config file or passed as arguments)
        string desiredTitle = "Quarterly Sales Overview";
        bool showLegend = false; // Set to true to display the legend

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Access the chart object.
        Chart chart = chartShape.Chart;

        // Update the chart title.
        ChartTitle title = chart.Title;
        title.Text = desiredTitle;
        title.Show = true; // Ensure the title is visible.

        // Toggle legend visibility based on the user preference.
        ChartLegend legend = chart.Legend;
        legend.Position = showLegend ? LegendPosition.Right : LegendPosition.None;

        // Save the document to the output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "updated-chart.docx");
        doc.Save(outputPath);
    }
}
