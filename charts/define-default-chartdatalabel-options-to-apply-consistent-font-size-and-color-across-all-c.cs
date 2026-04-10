using System;
using System.Drawing;
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
        if (!chartShape.HasChart)
            throw new InvalidOperationException("Inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Remove the demo series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Define categories and values for two series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Revenue", categories, new double[] { 15000, 21000, 18000, 24000 });
        chart.Series.Add("Expenses", categories, new double[] { 8000, 9500, 7000, 11000 });

        // Apply consistent data label settings to every series.
        foreach (ChartSeries series in chart.Series)
        {
            // Enable data labels for the series.
            series.HasDataLabels = true;

            // Show the value in each label.
            series.DataLabels.ShowValue = true;

            // Set a uniform font size and color for all labels in the series.
            series.DataLabels.Font.Size = 12;
            series.DataLabels.Font.Color = Color.DarkBlue;
        }

        // Save the document.
        doc.Save("ChartDataLabels_DefaultOptions.docx");
    }
}
