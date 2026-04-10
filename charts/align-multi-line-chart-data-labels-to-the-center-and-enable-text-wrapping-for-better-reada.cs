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

        // Add a custom series with categories and values.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120.5, 95.3, 143.8, 110.0 };
        ChartSeries series = chart.Series.Add("Sales", categories, values);

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Configure the data labels:
        // - Show category name, series name and value.
        // - Use a line break as separator to create multi‑line labels.
        // - Align the label to the center of the data point.
        ChartDataLabelCollection dataLabels = series.DataLabels;
        dataLabels.ShowCategoryName = true;
        dataLabels.ShowSeriesName = true;
        dataLabels.ShowValue = true;
        dataLabels.Separator = "\n";                     // Multi‑line.
        dataLabels.Position = ChartDataLabelPosition.Center; // Center alignment.

        // Save the document.
        doc.Save("AlignedMultiLineDataLabels.docx");
    }
}
