using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Clear default demo series and add custom data.
        chart.Series.Clear();
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        chart.Series.Add("Series 1", categories, new double[] { 10, 20, 30 });
        chart.Series.Add("Series 2", categories, new double[] { 15, 25, 35 });

        // Move the legend to the top right corner.
        ChartLegend legend = chart.Legend;
        legend.Position = LegendPosition.TopRight;

        // Set the legend background fill to light gray.
        legend.Format.Fill.Solid(Color.LightGray);

        // Save the document.
        doc.Save("ChartLegendExample.docx");
    }
}
