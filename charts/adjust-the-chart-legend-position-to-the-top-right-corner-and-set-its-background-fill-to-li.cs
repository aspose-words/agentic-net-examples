using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Ensure the chart has at least one series.
        chart.Series.Clear();
        chart.Series.Add("Series 1", new double[] { 10, 20, 30 });

        // Move the legend to the top‑right corner.
        chart.Legend.Position = LegendPosition.TopRight;

        // Set the legend background fill to light gray.
        chart.Legend.Format.Fill.ForeColor = Color.LightGray;

        // Save the document.
        doc.Save("ChartLegendModified.docx");
    }
}
