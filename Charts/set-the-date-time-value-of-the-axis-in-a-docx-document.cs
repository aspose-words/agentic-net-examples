using System;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added
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

        // Remove the demo series that Aspose adds by default.
        chart.Series.Clear();

        // Add a series that uses DateTime values for the X‑axis.
        chart.Series.Add(
            "Sample Series",
            new[]
            {
                new DateTime(2023, 1, 1),
                new DateTime(2023, 1, 5),
                new DateTime(2023, 1, 10)
            },
            new[] { 10.0, 20.0, 15.0 });

        // Configure the X‑axis to display date/time values.
        ChartAxis xAxis = chart.AxisX;

        // Set explicit lower and upper bounds for the axis.
        xAxis.Scaling.Minimum = new AxisBound(new DateTime(2022, 12, 31));
        xAxis.Scaling.Maximum = new AxisBound(new DateTime(2023, 1, 15));

        // Define the smallest time unit shown on the axis (days).
        xAxis.BaseTimeUnit = AxisTimeUnit.Days;

        // Set major and minor units (5 days and 1 day respectively).
        xAxis.MajorUnit = 5.0;
        xAxis.MinorUnit = 1.0;

        // Enable gridlines for better readability.
        xAxis.HasMajorGridlines = true;
        xAxis.HasMinorGridlines = true;

        // Save the document to a DOCX file.
        doc.Save("DateTimeAxis.docx");
    }
}
