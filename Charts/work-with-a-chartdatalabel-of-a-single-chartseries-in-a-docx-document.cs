using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added for Shape
using Aspose.Words.Drawing.Charts;   // Chart, ChartSeries, ChartDataLabel, etc.

class ChartDataLabelExample
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart.
        Shape chartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = chartShape.Chart;

        // Work with the first series.
        ChartSeries series = chart.Series[0];
        series.HasDataLabels = true; // Enable data labels for the series.

        // Access a specific data label (e.g., the first point).
        ChartDataLabel dataLabel = series.DataLabels[0];

        // Configure the data label.
        dataLabel.ShowValue = true;
        dataLabel.ShowCategoryName = true;
        dataLabel.Separator = " | ";
        dataLabel.Format.Fill.Color = Color.Green;

        // Uncomment to reset the label to default formatting.
        // dataLabel.ClearFormat();

        // Save the document.
        doc.Save("ChartDataLabelExample.docx");
    }
}
