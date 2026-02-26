using System;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added for Shape
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart.
        Shape shape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = shape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a series with DateTime values on the X‑axis.
        chart.Series.Add("Sample Series",
            new[]
            {
                new DateTime(2022, 1, 1),
                new DateTime(2022, 1, 5),
                new DateTime(2022, 1, 10)
            },
            new[] { 10.0, 20.0, 15.0 });

        // Configure the X‑axis to use date/time bounds.
        ChartAxis xAxis = chart.AxisX;
        xAxis.Scaling.Minimum = new AxisBound(new DateTime(2021, 12, 31));
        xAxis.Scaling.Maximum = new AxisBound(new DateTime(2022, 1, 15));
        xAxis.BaseTimeUnit = AxisTimeUnit.Days;   // Smallest unit displayed.
        xAxis.MajorUnit = 5.0;                    // Major tick every 5 days.
        xAxis.MinorUnit = 1.0;                    // Minor tick every day.

        // Save the document.
        doc.Save("DateTimeAxis.docx");
    }
}
