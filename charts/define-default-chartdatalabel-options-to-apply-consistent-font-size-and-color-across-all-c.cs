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

        // Define categories and series data.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Product A", categories, new double[] { 120, 150, 170, 130 });
        chart.Series.Add("Product B", categories, new double[] { 80, 110, 140, 100 });

        // Define the default font size and color for data labels.
        const float defaultFontSize = 10f;
        Color defaultFontColor = Color.White;

        // Apply the default data label settings to every series.
        foreach (ChartSeries series in chart.Series)
        {
            // Enable data labels for the series.
            series.HasDataLabels = true;

            // Show the value in each data label.
            series.DataLabels.ShowValue = true;

            // Set the default font size and color for all data labels in this series.
            series.DataLabels.Font.Size = defaultFontSize;
            series.DataLabels.Font.Color = defaultFontColor;
        }

        // Save the document.
        doc.Save("ChartDataLabelDefaults.docx");
    }
}
