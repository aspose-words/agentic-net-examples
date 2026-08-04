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

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add sample data to the chart.
        string[] categories = { "Category A", "Category B", "Category C" };
        chart.Series.Add("Sample Series", categories, new double[] { 10, 20, 30 });

        // Apply a semi‑transparent background fill to create a watermark effect.
        chart.Format.Fill.Solid(Color.LightGray);
        chart.Format.Fill.Transparency = 0.5; // 50 % transparency.

        // Save the document.
        doc.Save("ChartWatermark.docx");
    }
}
