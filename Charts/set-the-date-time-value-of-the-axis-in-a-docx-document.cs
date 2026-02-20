using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // <-- added
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart of size 500x300 points.
        Shape shape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = shape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a series where the X‑axis values are DateTime objects.
        chart.Series.Add(
            "Sample Series",
            new[]
            {
                new DateTime(2023, 1, 1),
                new DateTime(2023, 1, 5),
                new DateTime(2023, 1, 10)
            },
            new[] { 10.0, 20.0, 15.0 });

        // Access the X axis (the date/time axis).
        ChartAxis xAxis = chart.AxisX;

        // Set explicit lower and upper bounds for the axis.
        xAxis.Scaling.Minimum = new AxisBound(new DateTime(2022, 12, 31).ToOADate());
        xAxis.Scaling.Maximum = new AxisBound(new DateTime(2023, 1, 15).ToOADate());

        // Define the base time unit and the major/minor units.
        xAxis.BaseTimeUnit = AxisTimeUnit.Days;   // axis measured in days
        xAxis.MajorUnit = 5.0;                    // major tick every 5 days
        xAxis.MinorUnit = 1.0;                    // minor tick every 1 day

        // Optional visual tweaks.
        xAxis.MajorTickMark = AxisTickMark.Cross;
        xAxis.MinorTickMark = AxisTickMark.Outside;
        xAxis.HasMajorGridlines = true;
        xAxis.HasMinorGridlines = true;

        // Save the document to a DOCX file.
        doc.Save("DateTimeAxis.docx");
    }
}
