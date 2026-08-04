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

        // Insert a pie chart. Pie charts naturally display percentages.
        Shape chartShape = builder.InsertChart(ChartType.Pie, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Add a custom series with three categories.
        ChartSeries series = chart.Series.Add(
            "Sample Series",
            new[] { "Category A", "Category B", "Category C" },
            new[] { 30.0, 45.0, 25.0 });

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Configure each data label to show percentage with one decimal place.
        for (int i = 0; i < series.DataLabels.Count; i++)
        {
            ChartDataLabel label = series.DataLabels[i];
            label.ShowPercentage = true;                     // Show the percentage value.
            label.NumberFormat.FormatCode = "0.0%";          // One decimal place.
            label.ShowValue = false;                         // Hide the raw value (optional).
        }

        // Save the document.
        doc.Save("ChartDataLabelPercentage.docx");
    }
}
