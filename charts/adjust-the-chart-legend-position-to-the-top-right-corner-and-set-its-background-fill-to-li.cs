using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;          // Required for the Shape class
using Aspose.Words.Drawing.Charts;   // Chart‑related types

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);

        // Verify that the inserted shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Access the chart object.
        Chart chart = chartShape.Chart;

        // Move the legend to the top‑right corner.
        ChartLegend legend = chart.Legend;
        legend.Position = LegendPosition.TopRight;

        // Set the legend background fill to light gray.
        legend.Format.Fill.Solid(Color.LightGray);

        // Save the document.
        doc.Save("ChartLegendTopRight.docx");
    }
}
