using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

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

        // Clear default demo data and add custom series.
        chart.Series.Clear();
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        chart.Series.Add("Series 1", categories, new double[] { 10, 20, 30 });

        // Configure the legend border stroke.
        ChartLegend legend = chart.Legend;
        legend.Format.Stroke.Weight = 2.0;               // Thickness
        legend.Format.Stroke.DashStyle = DashStyle.Dash; // Dash style
        legend.Format.Stroke.Color = Color.DarkRed;      // Optional color

        // Save the document.
        doc.Save("ChartLegendBorder.docx");
    }
}
