using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added for Shape
using Aspose.Words.Drawing.Charts;

class ScatterChartExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will contain the chart.
        builder.Writeln("Scatter chart inserted below:");

        // Insert a scatter chart with the desired size (width, height) in points.
        // This uses the InsertChart(ChartType, double, double) overload.
        Shape chartShape = builder.InsertChart(ChartType.Scatter, 400, 300);

        // Get the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Add a series with X and Y values.
        chart.Series.Add(
            "Sample Series",
            new double[] { 1.0, 2.5, 4.0, 5.5 },   // X‑values
            new double[] { 2.0, 3.5, 1.0, 4.5 }    // Y‑values
        );

        // Save the document.
        doc.Save("ScatterChart.docx");
    }
}
