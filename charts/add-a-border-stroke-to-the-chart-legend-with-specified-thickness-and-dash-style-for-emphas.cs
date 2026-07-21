using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Optional: clear the demo data and add custom series.
        chart.Series.Clear();
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        chart.Series.Add("Series 1", categories, new double[] { 10, 20, 30 });
        chart.Series.Add("Series 2", categories, new double[] { 15, 25, 35 });

        // Configure the legend border: set thickness and dash style.
        ChartLegend legend = chart.Legend;
        legend.Format.Stroke.Weight = 2.0;               // Thickness of the border.
        legend.Format.Stroke.DashStyle = DashStyle.Dash; // Dash style.
        legend.Format.Stroke.Color = Color.DarkBlue;     // Border color (optional).

        // Save the document to the working directory.
        doc.Save("ChartLegendBorder.docx");
    }
}
