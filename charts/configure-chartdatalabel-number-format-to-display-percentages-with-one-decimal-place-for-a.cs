using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart. Pie charts support percentage data labels.
        Shape chartShape = builder.InsertChart(ChartType.Pie, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Add a custom series with categories and values.
        ChartSeries series = chart.Series.Add(
            "Sample Series",
            new[] { "Category A", "Category B", "Category C" },
            new[] { 30.0, 55.0, 15.0 });

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Configure each data label to show the percentage with one decimal place.
        foreach (ChartDataLabel dataLabel in series.DataLabels)
        {
            dataLabel.ShowPercentage = true;
            dataLabel.NumberFormat.FormatCode = "0.0%";
        }

        // Save the document.
        doc.Save("ChartDataLabelPercentage.docx");
    }
}
