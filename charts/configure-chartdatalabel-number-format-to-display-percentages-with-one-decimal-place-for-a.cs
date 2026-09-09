using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace ChartDataLabelPercentageExample
{
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

            // Remove the default demo series.
            chart.Series.Clear();

            // Add a custom series with categories and values.
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };
            double[] values = { 0.25, 0.35, 0.20, 0.20 };
            chart.Series.Add("Quarterly", categories, values);

            // Enable data labels for each series and set the number format to show percentages
            // with one decimal place.
            foreach (ChartSeries series in chart.Series)
            {
                series.HasDataLabels = true;
                ChartDataLabelCollection labels = series.DataLabels;
                labels.ShowValue = true;
                // Format code "0.0%" displays percentages with one decimal place.
                labels.NumberFormat.FormatCode = "0.0%";
            }

            // Save the document.
            doc.Save("ChartDataLabelPercentage.docx");
        }
    }
}
