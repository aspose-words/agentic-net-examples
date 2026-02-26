using System;
using Aspose.Words;
using Aspose.Words.Drawing; // <-- added for Shape
using Aspose.Words.Drawing.Charts;

class ScatterChartExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a scatter chart with a size of 500x500 pixels (converted to points).
        double width = ConvertUtil.PixelToPoint(500);
        double height = ConvertUtil.PixelToPoint(500);
        Shape chartShape = builder.InsertChart(ChartType.Scatter, width, height);

        // Access the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Add the first data series (X and Y values must have the same length).
        chart.Series.Add(
            "Series 1",
            new double[] { 3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6 },
            new double[] { 3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9 });

        // Add a second data series.
        chart.Series.Add(
            "Series 2",
            new double[] { 2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3 },
            new double[] { 7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6 });

        // Save the document containing the scatter chart.
        doc.Save("ScatterChart.docx");
    }
}
