using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class ChartDataLabelFontExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Ensure the chart has at least one series.
        if (chart.Series.Count == 0)
            throw new InvalidOperationException("The chart must contain at least one series.");

        // Work with the first series.
        ChartSeries series = chart.Series[0];

        // Enable data labels for the series.
        series.HasDataLabels = true;
        series.DataLabels.ShowValue = true;

        // Customize the font of all data labels in the series.
        series.DataLabels.Font.Name = "Arial";
        series.DataLabels.Font.Size = 14;
        series.DataLabels.Font.Bold = true;
        series.DataLabels.Font.Color = Color.DarkBlue;

        // Save the document.
        doc.Save("ChartDataLabelFont.docx");
    }
}
