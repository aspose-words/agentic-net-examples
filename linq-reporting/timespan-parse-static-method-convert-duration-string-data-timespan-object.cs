using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // <-- added for Shape
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

        // Example duration strings that we want to plot on the X‑axis.
        string[] durationStrings = { "01:15:00", "02:30:00", "03:45:00" };

        // For a Line chart the X‑values are expected to be strings (category labels).
        // Therefore we keep the original string array – no need for ChartXValue.
        string[] xValues = durationStrings;

        // Corresponding Y‑values for the series.
        double[] yValues = { 10.0, 20.0, 30.0 };

        // Add a new series that uses the string X‑values on the X‑axis.
        chart.Series.Add("Duration Series", xValues, yValues);

        // Optional: configure the X‑axis to display time values.
        ChartAxis xAxis = chart.AxisX;
        xAxis.BaseTimeUnit = AxisTimeUnit.Days; // Smallest unit shown on the axis.
        xAxis.MajorUnit = 1.0;                  // One day per major tick.
        xAxis.HasMajorGridlines = true;

        // Save the document containing the chart.
        doc.Save("ChartWithTimeSpan.docx");
    }
}
