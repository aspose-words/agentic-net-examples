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

        // Insert a line chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a series that uses DateTime values on the X‑axis.
        chart.Series.Add(
            "Sample Series",
            new[]
            {
                new DateTime(2023, 1, 1),
                new DateTime(2023, 1, 5),
                new DateTime(2023, 1, 10)
            },
            new[] { 10.0, 20.0, 15.0 });

        // Access the X axis (time category axis).
        ChartAxis xAxis = chart.AxisX;

        // Set explicit lower and upper bounds using AxisBound(DateTime).
        xAxis.Scaling.Minimum = new AxisBound(new DateTime(2022, 12, 31));
        xAxis.Scaling.Maximum = new AxisBound(new DateTime(2023, 1, 15));

        // Define the smallest time unit and major/minor units.
        xAxis.BaseTimeUnit = AxisTimeUnit.Days;   // axis displays days
        xAxis.MajorUnit = 5.0;                    // major tick every 5 days
        xAxis.MinorUnit = 1.0;                    // minor tick every day

        // Optional: show gridlines for better visibility.
        xAxis.HasMajorGridlines = true;
        xAxis.HasMinorGridlines = true;

        // Save the document with the configured date‑time axis.
        doc.Save("DateTimeAxis.docx");
    }
}
