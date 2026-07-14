using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ChartDataLabelFontExample
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Clear the default demo series and add a custom series.
        chart.Series.Clear();
        chart.Series.Add("Sales", new[] { "Q1", "Q2", "Q3" }, new double[] { 15000, 23000, 18000 });

        // Enable data labels for the series.
        ChartSeries series = chart.Series[0];
        series.HasDataLabels = true;

        // Customize the data label font: typeface, size, and bold styling.
        series.DataLabels.Font.Name = "Arial";
        series.DataLabels.Font.Size = 14;
        series.DataLabels.Font.Bold = true;

        // Save the document.
        doc.Save("ChartDataLabelFont.docx");
    }
}
