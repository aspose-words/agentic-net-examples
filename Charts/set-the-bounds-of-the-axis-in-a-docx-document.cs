using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a custom series with date/time values on the X‑axis.
        chart.Series.Add("Sample Series",
            new[]
            {
                new DateTime(2023, 1, 1),
                new DateTime(2023, 1, 5),
                new DateTime(2023, 1, 10)
            },
            new[] { 10.0, 20.0, 15.0 });

        // Set the X‑axis bounds (date axis uses OADate double values).
        ChartAxis xAxis = chart.AxisX;
        // Minimum bound.
        xAxis.Scaling.Minimum = new AxisBound(new DateTime(2022, 12, 25).ToOADate());
        // Maximum bound.
        xAxis.Scaling.Maximum = new AxisBound(new DateTime(2023, 1, 15).ToOADate());

        // Set the Y‑axis bounds.
        ChartAxis yAxis = chart.AxisY;
        yAxis.Scaling.Minimum = new AxisBound(0);
        yAxis.Scaling.Maximum = new AxisBound(30);

        // Save the document.
        doc.Save("AxisBounds.docx");
    }
}
