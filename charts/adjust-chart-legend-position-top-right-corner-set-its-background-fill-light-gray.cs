using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class ChartLegendExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart (column type) into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 450, 300);
        Chart chart = chartShape.Chart;

        // Access the chart's legend.
        ChartLegend legend = chart.Legend;

        // Move the legend to the top‑right corner.
        legend.Position = LegendPosition.TopRight;

        // Set the legend background fill to a light gray color.
        legend.Format.Fill.Solid(Color.LightGray);

        // Save the document.
        doc.Save("ChartLegendTopRight.docx");
    }
}
