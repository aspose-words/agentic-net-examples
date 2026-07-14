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

        // Clear default demo data and add custom series.
        chart.Series.Clear();
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        chart.Series.Add("Series 1", categories, new double[] { 10, 20, 30 });

        // Apply a border stroke to the chart legend.
        ChartLegend legend = chart.Legend;
        legend.Format.Stroke.Weight = 2.0;               // Thickness of the border.
        legend.Format.Stroke.DashStyle = DashStyle.Dash; // Dash style for emphasis.
        legend.Format.Stroke.Color = Color.DarkRed;      // Optional: set border color.

        // Save the document.
        doc.Save("ChartWithLegendBorder.docx");
    }
}
