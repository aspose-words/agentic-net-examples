using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class InsertScatterChart
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a scatter chart shape (width: 500 points, height: 300 points).
        Shape chartShape = builder.InsertChart(ChartType.Scatter, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Add the first data series – X and Y values must be of equal length.
        chart.Series.Add("Series 1",
            new[] { 3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6 },
            new[] { 3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9 });

        // Add a second data series.
        chart.Series.Add("Series 2",
            new[] { 2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3 },
            new[] { 7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6 });

        // Optionally set a chart title.
        chart.Title.Text = "Sample Scatter Chart";
        chart.Title.Show = true;

        // Save the document to a DOCX file.
        doc.Save("ScatterChart.docx");
    }
}
