using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Add a series with categories and values.
        string[] categories = { "Category A", "Category B", "Category C" };
        double[] values = { 120.5, 85.3, 97.8 };
        chart.Series.Add("Sample Series", categories, values);

        // Enable data labels for the series.
        ChartSeries series = chart.Series[0];
        series.HasDataLabels = true;

        // Align data labels to the center.
        series.DataLabels.Position = ChartDataLabelPosition.Center;

        // Use a line break as the separator to allow multi‑line labels (text wrapping).
        series.DataLabels.Separator = "\n";

        // Optionally show both category name and value.
        series.DataLabels.ShowCategoryName = true;
        series.DataLabels.ShowValue = true;

        // Save the document.
        doc.Save("AlignedWrappedDataLabels.docx");
    }
}
