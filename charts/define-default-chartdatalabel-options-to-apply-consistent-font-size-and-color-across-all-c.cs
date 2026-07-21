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
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Define categories and add two custom series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Product A", categories, new double[] { 120, 150, 170, 200 });
        chart.Series.Add("Product B", categories, new double[] { 80, 110, 130, 160 });

        // Apply default data label options to every series.
        foreach (ChartSeries series in chart.Series)
        {
            // Enable data labels for the series.
            series.HasDataLabels = true;

            // Set a consistent font size and color for all data labels in the series.
            series.DataLabels.Font.Size = 10;
            series.DataLabels.Font.Color = Color.DarkBlue;

            // Show the value in each label.
            series.DataLabels.ShowValue = true;
        }

        // Save the document.
        doc.Save("DefaultDataLabelOptions.docx");
    }
}
