using System;
using System.Drawing;
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

        // Define categories for the X‑axis.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };

        // Add two custom series with sample data.
        chart.Series.Add("Product A", categories, new double[] { 10, 20, 30, 40 });
        chart.Series.Add("Product B", categories, new double[] { 15, 25, 35, 45 });

        // Apply default data‑label settings to every series.
        foreach (ChartSeries series in chart.Series)
        {
            // Enable data labels for the series.
            series.HasDataLabels = true;

            // Show the value in each label.
            series.DataLabels.ShowValue = true;

            // Set a consistent font size and color for all labels in the series.
            series.DataLabels.Font.Size = 12;
            series.DataLabels.Font.Color = Color.DarkBlue;
        }

        // Save the document.
        doc.Save("ChartDataLabelDefaults.docx");
    }
}
