using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class ChartDataLabelExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a custom series with three data points.
        ChartSeries series = chart.Series.Add(
            "Sales",
            new[] { "Q1", "Q2", "Q3" },
            new[] { 120.5, 150.0, 180.75 });

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Access the first data label in the series (index 0).
        ChartDataLabel firstLabel = series.DataLabels[0];

        // Change the fill color of this data label to red.
        firstLabel.Format.Fill.Color = Color.Red;

        // Optionally clear the format of the second data label.
        ChartDataLabel secondLabel = series.DataLabels[1];
        secondLabel.ClearFormat();

        // Save the document to a DOCX file.
        doc.Save("ChartDataLabelExample.docx");
    }
}
