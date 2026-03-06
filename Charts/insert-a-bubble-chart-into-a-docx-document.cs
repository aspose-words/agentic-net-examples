using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bubble chart with the desired dimensions (width, height in points).
        Shape chartShape = builder.InsertChart(ChartType.Bubble, 500, 350);
        Chart chart = chartShape.Chart;

        // Remove the automatically generated demo series.
        chart.Series.Clear();

        // Add a custom series: X values, Y values, and bubble sizes.
        chart.Series.Add(
            "Sample Series",
            new double[] { 1.1, 5.0, 9.8 },   // X values
            new double[] { 1.2, 4.9, 9.9 },   // Y values
            new double[] { 2.0, 4.0, 8.0 }    // Bubble sizes
        );

        // Enable data labels for the series and display bubble sizes on the labels.
        ChartSeries series = chart.Series[0];
        series.HasDataLabels = true;
        series.DataLabels.ShowBubbleSize = true;

        // Save the document containing the bubble chart.
        doc.Save("BubbleChart.docx");
    }
}
