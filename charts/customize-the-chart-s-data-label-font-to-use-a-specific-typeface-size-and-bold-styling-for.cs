using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a custom series with categories and values.
        string[] categories = new[] { "Q1", "Q2", "Q3", "Q4" };
        double[] values = new double[] { 120, 150, 180, 200 };
        ChartSeries series = chart.Series.Add("Revenue", categories, values);

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Customize the data label font: typeface, size, and bold styling.
        series.DataLabels.Font.Name = "Arial";
        series.DataLabels.Font.Size = 14;
        series.DataLabels.Font.Bold = true;

        // Save the document.
        doc.Save("CustomDataLabelFont.docx");
    }
}
