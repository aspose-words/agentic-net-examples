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
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Define categories and values for two series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Product A", categories, new double[] { 10, 20, 30, 40 });
        chart.Series.Add("Product B", categories, new double[] { 15, 25, 35, 45 });

        // Enable data labels and configure them to show percentages with one decimal place.
        foreach (ChartSeries series in chart.Series)
        {
            series.HasDataLabels = true;

            // Iterate over each data label in the series.
            for (int i = 0; i < series.DataLabels.Count; i++)
            {
                ChartDataLabel label = series.DataLabels[i];
                label.ShowPercentage = true;
                label.NumberFormat.FormatCode = "0.0%";
            }
        }

        // Save the document.
        doc.Save("ChartDataLabelPercentage.docx");
    }
}
