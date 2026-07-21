using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // Required for the Shape class
using Aspose.Words.Drawing.Charts;   // Chart‑related APIs

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 350);
        Chart chart = chartShape.Chart;

        // Use the first automatically generated series.
        ChartSeries series = chart.Series[0];

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Customize the font of all data labels in this series.
        series.DataLabels.Font.Name = "Arial";
        series.DataLabels.Font.Size = 14;
        series.DataLabels.Font.Bold = true;

        // Save the document.
        doc.Save("ChartDataLabelFont.docx");
    }
}
