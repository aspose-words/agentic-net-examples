using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a specific size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the default series and add custom data.
        chart.Series.Clear();
        string[] categories = new string[] { "Category 1", "Category 2" };
        chart.Series.Add("Series 1", categories, new double[] { 1, 2 });
        chart.Series.Add("Series 2", categories, new double[] { 3, 4 });

        // ----- Fill formatting -----
        // Set the chart background to a solid dark slate gray color.
        chart.Format.Fill.Solid(Color.DarkSlateGray);

        // Set the chart title background fill.
        chart.Title.Format.Fill.Solid(Color.LightGoldenrodYellow);
        // Set the legend background fill.
        chart.Legend.Format.Fill.Solid(Color.LightGoldenrodYellow);

        // ----- Stroke (line) formatting -----
        // Set the chart outline (stroke) to a solid red line with a thickness of 2 points.
        chart.Format.Stroke.Color = Color.Red;
        chart.Format.Stroke.Weight = 2.0; // points
        chart.Format.Stroke.DashStyle = DashStyle.Solid;

        // Set the title outline.
        chart.Title.Format.Stroke.Color = Color.DarkGoldenrod;
        chart.Title.Format.Stroke.Weight = 1.0;
        // Set the legend outline.
        chart.Legend.Format.Stroke.Color = Color.DarkGoldenrod;
        chart.Legend.Format.Stroke.Weight = 1.0;

        // Save the document to a DOCX file.
        doc.Save("ChartWithFormatting.docx");
    }
}
