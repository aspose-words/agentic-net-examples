using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Required for the Shape class
using Aspose.Words.Drawing.Charts;        // Chart related classes

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with default demo data.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Retrieve the collection of series in the chart.
        ChartSeriesCollection seriesCollection = chart.Series;

        // Iterate through each series and modify its data points.
        foreach (ChartSeries series in seriesCollection)
        {
            // Preserve the original number of data points (categories).
            int originalPointCount = series.XValues.Count;

            // Remove existing values while keeping any formatting.
            series.ClearValues();

            // Add new data points. Here we use simple string categories and incremental values.
            for (int i = 0; i < originalPointCount; i++)
            {
                string category = $"Category {i + 1}";
                double newValue = (i + 1) * 10; // Example new Y value.

                series.Add(ChartXValue.FromString(category), ChartYValue.FromDouble(newValue));
            }
        }

        // Save the modified document. The chart will reflect the updated data when opened.
        doc.Save("ModifiedChart.docx");
    }
}
