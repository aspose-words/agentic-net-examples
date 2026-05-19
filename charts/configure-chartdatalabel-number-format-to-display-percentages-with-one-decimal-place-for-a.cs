using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Define categories and values (values are fractions to be shown as percentages).
        string[] categories = { "A", "B", "C" };
        double[] values1 = { 0.25, 0.50, 0.75 };
        double[] values2 = { 0.10, 0.30, 0.60 };

        // Add two series.
        ChartSeries series1 = chart.Series.Add("Series 1", categories, values1);
        ChartSeries series2 = chart.Series.Add("Series 2", categories, values2);

        // Enable data labels for each series and set the number format to one‑decimal‑place percentages.
        foreach (ChartSeries series in chart.Series)
        {
            series.HasDataLabels = true; // Show data labels.

            // Apply the format to every data label in the series.
            for (int i = 0; i < series.DataLabels.Count; i++)
            {
                series.DataLabels[i].NumberFormat.FormatCode = "0.0%";
                series.DataLabels[i].ShowValue = true; // Display the formatted value.
            }
        }

        // Save the document.
        doc.Save("ChartDataLabelPercentage.docx");
    }
}
