using System;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added
using Aspose.Words.Drawing.Charts;

class SetChartAxisBounds
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a scatter chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Scatter, 450, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a custom series with X and Y numeric values.
        chart.Series.Add("Series 1",
            new[] { 1.1, 5.4, 7.9, 3.5, 2.1, 9.7 },
            new[] { 2.1, 0.3, 0.6, 3.3, 1.4, 1.9 });

        // Set explicit minimum and maximum bounds for both axes (0 to 10).
        chart.AxisX.Scaling.Minimum = new AxisBound(0);
        chart.AxisX.Scaling.Maximum = new AxisBound(10);
        chart.AxisY.Scaling.Minimum = new AxisBound(0);
        chart.AxisY.Scaling.Maximum = new AxisBound(10);

        // Save the document to a DOCX file.
        doc.Save("AxisBounds.docx");
    }
}
