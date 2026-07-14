using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ChartDataLabelsThreshold
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Define categories and corresponding values.
        string[] categories = { "A", "B", "C", "D", "E" };
        double[] values = { 30, 70, 45, 90, 20 };

        // Add a custom series with the data.
        chart.Series.Add("Sample Series", categories, values);
        ChartSeries series = chart.Series[0];

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Threshold value – only points with a value greater than this will show a label.
        double threshold = 50.0;

        // Iterate over each data point and hide the label if the value does not exceed the threshold.
        for (int i = 0; i < series.YValues.Count; i++)
        {
            double pointValue = series.YValues[i].DoubleValue;
            ChartDataLabel dataLabel = series.DataLabels[i];

            if (pointValue <= threshold)
            {
                // Hide the data label for points below or equal to the threshold.
                dataLabel.IsHidden = true;
            }
            else
            {
                // Ensure the label is visible for points above the threshold.
                dataLabel.IsHidden = false;
                dataLabel.ShowValue = true;
            }
        }

        // Save the document with the configured chart.
        doc.Save("ChartDataLabelsThreshold.docx");
    }
}
