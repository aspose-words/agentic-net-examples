using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;          // Required for the Shape class
using Aspose.Words.Drawing.Charts;

public class ChartDataLabelDefaultsExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart and obtain its Chart object.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the demo data series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Define categories for the X‑axis.
        string[] categories = new[] { "Category 1", "Category 2", "Category 3" };

        // Add two series with sample values.
        chart.Series.Add("Series 1", categories, new double[] { 10, 20, 30 });
        chart.Series.Add("Series 2", categories, new double[] { 15, 25, 35 });

        // Apply default data‑label formatting to every series.
        foreach (ChartSeries series in chart.Series)
        {
            // Enable data labels for the series.
            series.HasDataLabels = true;

            // Set a consistent font size and color for all data labels in the series.
            series.DataLabels.Font.Size = 12;
            series.DataLabels.Font.Color = Color.DarkBlue;
        }

        // Save the document.
        doc.Save("ChartDataLabelDefaults.docx");
    }
}
