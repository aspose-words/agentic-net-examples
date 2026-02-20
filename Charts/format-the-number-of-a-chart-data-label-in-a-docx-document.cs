using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace FormatChartDataLabels
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new document and a DocumentBuilder to work with it.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart into the document.
            Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
            Chart chart = chartShape.Chart;

            // Remove the default demo series.
            chart.Series.Clear();

            // Add a custom series with sample categories and values.
            ChartSeries series = chart.Series.Add(
                "Sample Series",
                new[] { "A", "B", "C", "D" },
                new double[] { 12345, 67890, 23456, 78901 });

            // Enable data labels for the series.
            series.HasDataLabels = true;

            // Set the number format for all data labels of the series.
            // This uses the ChartNumberFormat class accessed via the DataLabels collection.
            series.DataLabels.NumberFormat.FormatCode = "#,##0";

            // Optionally, you can also set the format for an individual label:
            // series.DataLabels[0].NumberFormat.FormatCode = "#,##0";

            // Save the document.
            doc.Save("FormattedChartDataLabels.docx");
        }
    }
}
