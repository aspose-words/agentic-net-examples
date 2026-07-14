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

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Move the legend to the top right corner.
        ChartLegend legend = chart.Legend;
        legend.Position = LegendPosition.TopRight;

        // Set the legend background fill to light gray.
        legend.Format.Fill.Solid(Color.LightGray);

        // Save the document.
        doc.Save("ChartLegendTopRight.docx");
    }
}
