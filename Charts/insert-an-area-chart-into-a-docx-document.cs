using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

class InsertAreaChart
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an Area chart with a width of 500 points and a height of 300 points.
        Shape chartShape = builder.InsertChart(ChartType.Area, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Define X‑axis categories (dates) and Y‑axis values.
        DateTime[] dates = {
            new DateTime(2014, 3, 31),
            new DateTime(2017, 1, 23),
            new DateTime(2017, 6, 18),
            new DateTime(2019, 11, 22),
            new DateTime(2020, 9, 7)
        };

        double[] values = { 15.8, 21.5, 22.9, 28.7, 33.1 };

        // Add a series to the chart using the dates and values.
        chart.Series.Add("Series 1", dates, values);

        // Optionally set a title for the chart.
        chart.Title.Text = "Sample Area Chart";
        chart.Title.Show = true;

        // Save the document to a DOCX file.
        doc.Save("AreaChart.docx");
    }
}
