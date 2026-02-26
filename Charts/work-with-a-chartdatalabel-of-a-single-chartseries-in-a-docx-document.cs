using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added for Shape
using Aspose.Words.Drawing.Charts;   // <-- for Chart, ChartSeries, ChartDataLabel

class ChartDataLabelExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 400);
        Chart chart = chartShape.Chart;

        // Work with the first series in the chart.
        ChartSeries series = chart.Series[0];

        // Enable data labels for this series.
        series.HasDataLabels = true;

        // Access a specific data label (second point, index 1).
        ChartDataLabel dataLabel = series.DataLabels[1];

        // Change the fill color of the data label.
        dataLabel.Format.Fill.Color = Color.Blue;

        // Set a custom separator string for the label.
        dataLabel.Separator = " | ";

        // Clear the label's format, reverting to defaults.
        dataLabel.ClearFormat();

        // Save the document to a DOCX file.
        doc.Save("ChartDataLabelExample.docx");
    }
}
