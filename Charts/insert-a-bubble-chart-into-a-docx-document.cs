using System;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added
using Aspose.Words.Drawing.Charts;

class InsertBubbleChart
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a 2‑D bubble chart with the desired size (width and height are in points).
        // ChartType.Bubble specifies a standard bubble chart.
        Shape chartShape = builder.InsertChart(ChartType.Bubble, 500, 300);

        // Get the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Add a custom series with X values, Y values and bubble sizes (diameters).
        // The three arrays must be of equal length.
        ChartSeries series = chart.Series.Add(
            "Sample Series",
            new double[] { 1.1, 5.0, 9.8 },      // X‑values
            new double[] { 1.2, 4.9, 9.9 },      // Y‑values
            new double[] { 2.0, 4.0, 8.0 }       // Bubble sizes
        );

        // Enable data labels for the series and configure them to show bubble size,
        // category name and series name.
        series.HasDataLabels = true;
        series.DataLabels.ShowBubbleSize = true;
        series.DataLabels.ShowCategoryName = true;
        series.DataLabels.ShowSeriesName = true;

        // Optionally, adjust the overall bubble scale (percentage of default size).
        // This property applies to the series group of the bubble chart.
        chart.SeriesGroups[0].BubbleScale = 150; // 150 % size

        // Save the document to a DOCX file.
        doc.Save("BubbleChart.docx");
    }
}
