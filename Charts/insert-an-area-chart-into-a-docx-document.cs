using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added for Shape
using Aspose.Words.Drawing.Charts;   // Chart, ChartType, etc.

class InsertAreaChart
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an Area chart with the desired size (width and height in points).
        // 500 points ≈ 7 inches, 300 points ≈ 4.2 inches.
        Shape chartShape = builder.InsertChart(ChartType.Area, 500, 300);

        // Get the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Optional: set a title for the chart.
        chart.Title.Text = "Sample Area Chart";
        chart.Title.Show = true;

        // Remove the default demo series that Aspose.Words inserts.
        chart.Series.Clear();

        // Define X‑axis categories as dates.
        DateTime[] dates = {
            new DateTime(2014, 3, 31),
            new DateTime(2017, 1, 23),
            new DateTime(2017, 6, 18),
            new DateTime(2019, 11, 22),
            new DateTime(2020, 9, 7)
        };

        // Define Y‑axis values corresponding to each date.
        double[] values = { 15.8, 21.5, 22.9, 28.7, 33.1 };

        // Add a new series to the chart using the dates and values.
        chart.Series.Add("Series 1", dates, values);

        // Save the document to a DOCX file.
        doc.Save("AreaChart.docx");
    }
}
