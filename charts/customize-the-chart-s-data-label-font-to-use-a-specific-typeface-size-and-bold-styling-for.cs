using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace ChartDataLabelFontExample
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

            // Remove the demo data series that Aspose.Words adds by default.
            chart.Series.Clear();

            // Add a custom series with categories and values.
            chart.Series.Add(
                "Quarterly Sales",
                new[] { "Q1", "Q2", "Q3", "Q4" },
                new double[] { 15000, 23000, 18000, 27000 });

            // Get the first (and only) series we just added.
            ChartSeries series = chart.Series[0];

            // Enable data labels for the series.
            series.HasDataLabels = true;

            // Customize the font of all data labels in the series.
            series.DataLabels.Font.Name = "Calibri";
            series.DataLabels.Font.Size = 12;
            series.DataLabels.Font.Bold = true;

            // Save the document to a file.
            doc.Save("ChartDataLabelFont.docx");
        }
    }
}
