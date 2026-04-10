using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for the Shape class
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart.
        Shape chartShape = builder.InsertChart(ChartType.Pie, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Add a custom series with categories and values.
        chart.Series.Add("Sample Series",
            new[] { "Category A", "Category B", "Category C" },
            new[] { 30.0, 45.0, 25.0 });

        // Enable data labels and configure them to show percentages with one decimal place.
        foreach (ChartSeries series in chart.Series)
        {
            series.HasDataLabels = true;

            // Apply the setting to the whole series (simpler than per‑point).
            series.DataLabels.ShowPercentage = true;
            series.DataLabels.NumberFormat.FormatCode = "0.0%";
        }

        // Save the document.
        doc.Save("ChartDataLabelPercentage.docx");
    }
}
