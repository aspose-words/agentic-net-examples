using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

class ChartDataPointExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Line, 500, 350);
        Chart chart = chartShape.Chart;

        // Access the first series of the chart.
        ChartSeries series = chart.Series[0];

        // Retrieve the first data point of the series.
        ChartDataPoint dataPoint = series.DataPoints[0];

        // Set a custom fill color for this data point.
        dataPoint.Format.Fill.Color = Color.Green;

        // Change the marker symbol and size for the data point.
        dataPoint.Marker.Symbol = MarkerSymbol.Star;
        dataPoint.Marker.Size = 12;

        // Uncomment the following line to clear the custom formatting
        // and revert the data point to the series default format.
        // dataPoint.ClearFormat();

        // Save the document to a DOCX file.
        doc.Save("ChartDataPointExample.docx");
    }
}
