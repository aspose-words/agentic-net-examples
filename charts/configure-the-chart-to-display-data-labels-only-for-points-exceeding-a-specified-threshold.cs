using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Define categories and values.
        string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
        double[] values = { 120, 85, 150, 60, 200 };
        double threshold = 100; // Labels will be shown only for values > 100.

        // Add a custom series.
        chart.Series.Add("Sales", categories, values);
        ChartSeries series = chart.Series[0];

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Configure each data label.
        for (int i = 0; i < series.DataLabels.Count; i++)
        {
            ChartDataLabel label = series.DataLabels[i];
            // Show the value for every point.
            label.ShowValue = true;

            // Hide the label if the point's value does not exceed the threshold.
            double pointValue = series.YValues[i].DoubleValue;
            label.IsHidden = pointValue <= threshold;
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartWithConditionalLabels.docx");
        doc.Save(outputPath);
    }
}
