using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);

        // Ensure the inserted shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Add a series with categories that contain long text (simulating multi‑line labels).
        string[] categories = {
            "Very Long Category Name 1",
            "Very Long Category Name 2",
            "Very Long Category Name 3",
            "Very Long Category Name 4"
        };
        double[] values = { 10, 20, 30, 25 };
        chart.Series.Add("Sample Series", categories, values);

        // Enable data labels for the series.
        ChartSeries series = chart.Series[0];
        series.HasDataLabels = true;

        // Align each data label to the center and set a modest font size.
        // Aspose.Words automatically wraps text that exceeds the label bounds.
        foreach (ChartDataLabel label in series.DataLabels)
        {
            label.Position = ChartDataLabelPosition.Center;
            label.Font.Size = 8;
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlignedDataLabels.docx");
        doc.Save(outputPath);
    }
}
