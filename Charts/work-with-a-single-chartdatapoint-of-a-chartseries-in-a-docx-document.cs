using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

namespace ChartDataPointExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a line chart into the document.
            // Width = 500 points, Height = 350 points.
            Shape chartShape = builder.InsertChart(ChartType.Line, 500, 350);
            Chart chart = chartShape.Chart;

            // Access the second series (index 1) and its third data point (index 2).
            ChartSeries series = chart.Series[1];
            ChartDataPoint dataPoint = series.DataPoints[2];

            // Change the fill color of this data point to red.
            dataPoint.Format.Fill.ForeColor = Color.Red;

            // Set a custom marker for the data point: a diamond with size 12.
            dataPoint.Marker.Symbol = MarkerSymbol.Diamond;
            dataPoint.Marker.Size = 12;

            // Optionally clear the formatting to revert to series defaults.
            // dataPoint.ClearFormat();

            // Save the document to a DOCX file.
            string outputPath = "ChartDataPointExample.docx";
            doc.Save(outputPath);
        }
    }
}
