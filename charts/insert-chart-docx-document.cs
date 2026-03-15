using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart with a size of 300x300 pixels.
        // ConvertUtil.PixelToPoint converts pixel dimensions to points (the unit used by Word).
        Shape chartShape = builder.InsertChart(
            ChartType.Pie,
            ConvertUtil.PixelToPoint(300),
            ConvertUtil.PixelToPoint(300));

        // Retrieve the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Add a series with categories (fruit names) and corresponding values.
        chart.Series.Add(
            "My fruit",
            new[] { "Apples", "Bananas", "Cherries" },
            new[] { 1.3, 2.2, 1.5 });

        // Optional: set a visible title for the chart.
        chart.Title.Text = "Fruit Distribution";
        chart.Title.Show = true;
        chart.Title.Font.Size = 14;
        chart.Title.Font.Color = Color.Blue;

        // Save the document to a DOCX file.
        doc.Save("ChartDocument.docx");
    }
}
