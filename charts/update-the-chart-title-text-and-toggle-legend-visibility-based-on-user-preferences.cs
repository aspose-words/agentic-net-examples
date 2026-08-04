using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // User preferences
        string desiredTitle = "Sales Overview";
        bool showLegend = false; // Set to true to display the legend

        // Create a new document and insert a chart
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // Ensure the shape actually contains a chart
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Update the chart title
        ChartTitle title = chart.Title;
        title.Text = desiredTitle;
        title.Show = true; // Make sure the title is visible

        // Toggle legend visibility based on the preference
        ChartLegend legend = chart.Legend;
        legend.Position = showLegend ? LegendPosition.Right : LegendPosition.None;

        // Save the document
        doc.Save("UpdatedChart.docx");
    }
}
