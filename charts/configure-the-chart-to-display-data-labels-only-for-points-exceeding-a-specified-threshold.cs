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

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Define categories (X‑axis) and corresponding numeric values (Y‑axis).
        string[] categories = { "A", "B", "C", "D", "E" };
        double[] values = { 30, 60, 45, 80, 20 };

        // Add a single series with the above data.
        chart.Series.Add("Sample Series", categories, values);

        // Retrieve the series we just added.
        ChartSeries series = chart.Series[0];

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Threshold – only points with a value greater than this will show a label.
        double threshold = 50.0;

        // Configure each data label individually.
        for (int i = 0; i < series.DataLabels.Count; i++)
        {
            ChartDataLabel label = series.DataLabels[i];

            if (values[i] > threshold)
            {
                // Show the value (and optionally the category name) for points above the threshold.
                label.ShowValue = true;
                label.ShowCategoryName = true;
                label.IsHidden = false; // Ensure the label is not hidden.
            }
            else
            {
                // Hide the label for points that do not meet the threshold.
                label.ShowValue = false;
                label.ShowCategoryName = false;
                label.IsHidden = true;
            }
        }

        // Save the document containing the configured chart.
        doc.Save("ChartDataLabelsThreshold.docx");
    }
}
