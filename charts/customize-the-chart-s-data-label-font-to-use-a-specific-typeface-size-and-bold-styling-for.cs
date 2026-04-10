using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a defined size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a custom series with categories and values.
        string[] categories = new[] { "Q1", "Q2", "Q3", "Q4" };
        double[] values = new[] { 120.5, 150.0, 180.75, 200.25 };
        ChartSeries series = chart.Series.Add("Revenue", categories, values);

        // Enable data labels for the series.
        series.HasDataLabels = true;
        series.DataLabels.ShowValue = true; // Show the numeric value.

        // Customize the font of all data labels in this series.
        series.DataLabels.Font.Name = "Arial";
        series.DataLabels.Font.Size = 14;
        series.DataLabels.Font.Bold = true;
        series.DataLabels.Font.Color = Color.DarkBlue;

        // Optionally, customize individual data label (demonstration).
        // series.DataLabels[0].Font.Italic = true;

        // Save the document to the local file system.
        doc.Save("ChartDataLabelFontCustomization.docx");
    }
}
