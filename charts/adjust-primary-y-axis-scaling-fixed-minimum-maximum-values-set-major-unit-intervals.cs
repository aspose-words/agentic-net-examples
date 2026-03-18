using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // <-- added
using Aspose.Words.Drawing.Charts;

class YAxisScalingExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Clear any demo data series that come with the chart.
        chart.Series.Clear();

        // Add a simple data series for demonstration purposes.
        chart.Series.Add("Sample Series",
            new[] { "A", "B", "C", "D" },
            new double[] { 10, 30, 20, 40 });

        // Access the primary Y axis.
        ChartAxis yAxis = chart.AxisY;

        // Set fixed minimum and maximum values for the Y axis.
        yAxis.Scaling.Minimum = new AxisBound(0);   // Minimum value = 0
        yAxis.Scaling.Maximum = new AxisBound(50);  // Maximum value = 50

        // Define the major unit interval (distance between major tick marks).
        yAxis.MajorUnit = 10; // Major tick every 10 units

        // Optional: adjust minor unit if desired.
        yAxis.MinorUnit = 5;  // Minor tick every 5 units

        // Save the document to disk.
        doc.Save("YAxisScaling.docx");
    }
}
