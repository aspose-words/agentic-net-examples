using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove default series and add custom data.
        chart.Series.Clear();
        string[] categories = new[] { "Category 1", "Category 2" };
        chart.Series.Add("Series 1", categories, new double[] { 1, 2 });
        chart.Series.Add("Series 2", categories, new double[] { 3, 4 });

        // Set chart background fill to a solid color.
        chart.Format.Fill.Solid(Color.DarkSlateGray);

        // Set chart outline (stroke) color and weight.
        chart.Format.Stroke.Color = Color.Red;
        chart.Format.Stroke.Weight = 2.0; // points

        // Save the document.
        doc.Save("ChartWithFormatting.docx");
    }
}
