using Aspose.Words;
using Aspose.Words.Drawing;               // <-- added for Shape
using Aspose.Words.Drawing.Charts;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Line, 500, 350);
        Chart chart = chartShape.Chart;

        // Access the second series (index 1) and its third data point (index 2).
        ChartSeries series = chart.Series[1];
        ChartDataPoint dataPoint = series.DataPoints[2];

        // Change the fill color of the selected data point.
        dataPoint.Format.Fill.Color = Color.Red;

        // Clear the formatting of the data point, reverting it to the series defaults.
        dataPoint.ClearFormat();

        // Save the document to a DOCX file.
        doc.Save("ChartDataPointExample.docx");
    }
}
