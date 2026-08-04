using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class AlignChartDataLabels
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Clear the default demo series.
        chart.Series.Clear();

        // Add a series with categories and values.
        string[] categories = { "Category A", "Category B", "Category C" };
        double[] values = { 120.5, 85.3, 97.8 };
        chart.Series.Add("Sample Series", categories, values);

        // Enable data labels for the series.
        ChartSeries series = chart.Series[0];
        series.HasDataLabels = true;

        // Align all data labels to the center of the data point.
        series.DataLabels.Position = ChartDataLabelPosition.Center;

        // Set font properties to help with wrapping.
        series.DataLabels.Font.Size = 10;
        series.DataLabels.Font.Name = "Arial";

        // Show category name and value to create multi‑line labels.
        series.DataLabels.ShowCategoryName = true;
        series.DataLabels.ShowValue = true;

        // Save the document.
        doc.Save("AlignedDataLabels.docx");
    }
}
