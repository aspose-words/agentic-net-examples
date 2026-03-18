using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added
using Aspose.Words.Drawing.Charts;

class ChartSeriesAddExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Column chart of size 500x300 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that comes with a newly inserted chart.
        chart.Series.Clear();

        // Define category names (X‑axis) and corresponding values (Y‑axis).
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120.5, 135.0, 150.2, 165.8 };

        // Add a labeled series in a single call using the overload that accepts a name,
        // an array of categories and an array of values.
        ChartSeries series = chart.Series.Add("Fiscal Year 2023", categories, values);

        // Optional: enable data labels for the series.
        series.HasDataLabels = true;

        // Ensure the output directory exists.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Save the document.
        doc.Save(Path.Combine(artifactsDir, "ChartSeriesAdd.docx"));
    }
}
