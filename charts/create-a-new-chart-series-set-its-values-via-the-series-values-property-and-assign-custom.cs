using System;
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

        // Remove the demo data series that come with a new chart.
        chart.Series.Clear();

        // Add a custom series with category labels (X‑axis) and numeric values (Y‑axis).
        chart.Series.Add(
            "Custom Series",
            new[] { "Category A", "Category B", "Category C" },
            new[] { 12.5, 23.0, 7.8 });

        // Save the document containing the customized chart.
        doc.Save("ChartSeriesValues.docx");
    }
}
