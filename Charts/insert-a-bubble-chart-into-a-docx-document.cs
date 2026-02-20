using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class InsertBubbleChart
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bubble chart shape with the desired size.
        Shape chartShape = builder.InsertChart(ChartType.Bubble, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series that Aspose.Words adds.
        chart.Series.Clear();

        // Add a bubble series: X values, Y values, and bubble sizes.
        chart.Series.Add(
            "Series 1",
            new double[] { 1.1, 5.0, 9.8 },   // X values
            new double[] { 1.2, 4.9, 9.9 },   // Y values
            new double[] { 2.0, 4.0, 8.0 }    // Bubble sizes
        );

        // Save the document containing the bubble chart.
        doc.Save("BubbleChart.docx");
    }
}
