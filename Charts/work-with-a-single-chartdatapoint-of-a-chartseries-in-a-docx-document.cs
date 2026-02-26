using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class ChartDataPointExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 350);
        Chart chart = chartShape.Chart;

        // Ensure the chart has at least one series.
        if (chart.Series.Count == 0)
            throw new InvalidOperationException("The chart does not contain any series.");

        // Get the first series in the chart.
        ChartSeries series = chart.Series[0];

        // Ensure the series has at least one data point.
        if (series.DataPoints.Count == 0)
            throw new InvalidOperationException("The series does not contain any data points.");

        // Access the first data point of the series.
        ChartDataPoint dataPoint = series.DataPoints[0];

        // Change the fill color of the data point to red.
        dataPoint.Format.Fill.Color = Color.Red;

        // Optionally clear the formatting of the data point, reverting to series defaults.
        dataPoint.ClearFormat();

        // Save the document to a DOCX file.
        doc.Save("ChartDataPointExample.docx");
    }
}
