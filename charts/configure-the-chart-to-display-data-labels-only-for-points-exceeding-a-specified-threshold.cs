using System;
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
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Define categories and values.
        string[] categories = { "A", "B", "C", "D", "E" };
        double[] values = { 5, 12, 8, 20, 3 };
        double threshold = 10.0;

        // Add a custom series.
        chart.Series.Add("Series 1", categories, values);
        ChartSeries series = chart.Series[0];

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Show data labels only for points whose value exceeds the threshold.
        for (int i = 0; i < series.YValues.Count; i++)
        {
            double pointValue = series.YValues[i].DoubleValue;

            // Hide all label components by default.
            series.DataLabels[i].ShowValue = false;
            series.DataLabels[i].ShowCategoryName = false;
            series.DataLabels[i].ShowSeriesName = false;

            if (pointValue > threshold)
            {
                // Show the value for points above the threshold.
                series.DataLabels[i].ShowValue = true;
            }
        }

        // Save the document.
        doc.Save("ChartWithConditionalLabels.docx");
    }
}
