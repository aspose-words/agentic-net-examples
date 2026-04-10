using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace AsposeChartsExample
{
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

            // Remove the demo data that comes with a newly inserted chart.
            chart.Series.Clear();

            // Define custom category labels and corresponding values.
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };
            double[] values = { 1600, 2100, 1900, 2300 };

            // Add a new series with the custom categories and values.
            // This overload automatically aligns categories (X) with values (Y).
            ChartSeries series = chart.Series.Add("Revenue", categories, values);

            // If you need to modify the series after creation, clear existing data
            // while preserving formatting, then add points individually.
            // series.ClearValues();
            // for (int i = 0; i < categories.Length; i++)
            // {
            //     series.Add(ChartXValue.FromString(categories[i]), ChartYValue.FromDouble(values[i]));
            // }

            // Save the document to the local file system.
            doc.Save("ChartSeriesValues.docx");
        }
    }
}
