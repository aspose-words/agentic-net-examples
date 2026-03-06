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

        // Insert a scatter chart with a specific size (width and height in points).
        double width = ConvertUtil.PixelToPoint(500);   // 500 pixels → points
        double height = ConvertUtil.PixelToPoint(400);  // 400 pixels → points
        Shape chartShape = builder.InsertChart(ChartType.Scatter, width, height);
        Chart chart = chartShape.Chart;

        // Remove the automatically generated demo series.
        chart.Series.Clear();

        // Add the first data series.
        chart.Series.Add(
            "Series 1",
            new double[] { 1, 2, 3, 4, 5 },          // X‑values
            new double[] { 2, 4, 1, 3, 5 }           // Y‑values
        );

        // Add the second data series.
        chart.Series.Add(
            "Series 2",
            new double[] { 1, 2, 3, 4, 5 },          // X‑values
            new double[] { 5, 3, 4, 2, 1 }           // Y‑values
        );

        // Optional: set a visible title for the chart.
        chart.Title.Text = "Sample Scatter Chart";
        chart.Title.Show = true;

        // Save the document to a DOCX file.
        doc.Save("ScatterChart.docx");
    }
}
