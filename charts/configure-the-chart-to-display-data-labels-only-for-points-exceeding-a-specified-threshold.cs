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

        // Define categories and corresponding values.
        string[] categories = { "A", "B", "C", "D", "E" };
        double[] values = { 12.5, 7.3, 15.8, 4.2, 9.0 };

        // Add a single series with the data.
        chart.Series.Add("Sample Series", categories, values);
        ChartSeries series = chart.Series[0];

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Threshold above which a data label will be shown.
        double threshold = 10.0;

        // Configure each data label: show only if the point's value exceeds the threshold.
        for (int i = 0; i < series.YValues.Count; i++)
        {
            double pointValue = series.YValues[i].DoubleValue;

            // Show the value label only for points exceeding the threshold.
            series.DataLabels[i].ShowValue = pointValue > threshold;

            // Hide the label completely for points below the threshold.
            series.DataLabels[i].IsHidden = pointValue <= threshold;
        }

        // Save the document.
        doc.Save("ChartWithConditionalDataLabels.docx");
    }
}
