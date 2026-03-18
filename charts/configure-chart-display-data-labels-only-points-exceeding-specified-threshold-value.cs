using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // <-- added
using Aspose.Words.Drawing.Charts;
using System.Drawing;

class ChartDataLabelThresholdExample
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 350);
        Chart chart = chartShape.Chart;

        // Clear the demo series.
        chart.Series.Clear();

        // Define categories and values.
        string[] categories = new[] { "Jan", "Feb", "Mar", "Apr", "May" };
        double[] values = new[] { 12.5, 7.3, 15.8, 4.2, 9.6 };

        // Add a custom series.
        ChartSeries series = chart.Series.Add("Sales", categories, values);

        // Enable data labels for the series.
        series.HasDataLabels = true;
        ChartDataLabelCollection dataLabels = series.DataLabels;
        dataLabels.ShowValue = true; // Show values where labels are visible.

        // Threshold value – only points greater than this will display a label.
        double threshold = 10.0;

        // Iterate through each data point and hide/show its label based on the threshold.
        for (int i = 0; i < series.YValues.Count; i++)
        {
            // Access the label for the current point.
            ChartDataLabel label = dataLabels[i];

            // Show the label only if the point's value exceeds the threshold.
            if (series.YValues[i].DoubleValue > threshold)
            {
                label.IsHidden = false; // Ensure the label is visible.
            }
            else
            {
                label.IsHidden = true; // Hide the label for values below or equal to the threshold.
            }
        }

        // Save the document.
        doc.Save("ChartDataLabelsThreshold.docx");
    }
}
