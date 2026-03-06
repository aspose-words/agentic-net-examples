using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added for Shape
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a scatter chart of size 450x300 points.
        Shape chartShape = builder.InsertChart(ChartType.Scatter, 450, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a custom series with sample X and Y values.
        chart.Series.Add("Series 1",
            new[] { 1.1, 5.4, 7.9, 3.5, 2.1, 9.7 },
            new[] { 2.1, 0.3, 0.6, 3.3, 1.4, 1.9 });

        // Set explicit bounds for the X‑axis (0 to 10).
        chart.AxisX.Scaling.Minimum = new AxisBound(0);
        chart.AxisX.Scaling.Maximum = new AxisBound(10);

        // Set explicit bounds for the Y‑axis (0 to 10).
        chart.AxisY.Scaling.Minimum = new AxisBound(0);
        chart.AxisY.Scaling.Maximum = new AxisBound(10);

        // Save the document containing the chart with custom axis bounds.
        doc.Save("AxisBounds.docx");
    }
}
