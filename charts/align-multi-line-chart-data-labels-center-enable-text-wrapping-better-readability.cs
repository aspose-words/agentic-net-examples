using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart.
        Shape chartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series.
        chart.Series.Clear();

        // Add a custom series.
        ChartSeries series = chart.Series.Add(
            "Series 1",
            new[] { "Jan", "Feb", "Mar", "Apr" },
            new[] { 10.0, 20.0, 15.0, 25.0 });

        // Enable data labels for the series.
        series.HasDataLabels = true;
        ChartDataLabelCollection dataLabels = series.DataLabels;

        // Show both category name and value.
        dataLabels.ShowCategoryName = true;
        dataLabels.ShowValue = true;

        // Align the data labels to the centre of the marker.
        dataLabels.Position = ChartDataLabelPosition.Center;

        // Use a line‑break as separator to allow multi‑line text (wrapping).
        dataLabels.Separator = "\n";

        // Optional: set a readable font.
        dataLabels.Font.Size = 10;
        dataLabels.Font.Color = Color.Black;

        // Save the document.
        doc.Save("ChartDataLabelsCenterWrap.docx");
    }
}
