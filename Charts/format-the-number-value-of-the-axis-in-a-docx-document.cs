using System;
using Aspose.Words;
using Aspose.Words.Drawing; // <-- added
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape shape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = shape.Chart;

        // Remove the default demo series to start with a clean chart.
        chart.Series.Clear();

        // Add a custom data series with categories and numeric values.
        chart.Series.Add(
            "Sample Series",
            new[] { "Category A", "Category B", "Category C", "Category D" },
            new double[] { 12345, 67890, 23456, 78901 });

        // Format the Y‑axis tick labels: use a thousand‑separator format and
        // disable linking to the source cell so the custom format is applied.
        chart.AxisY.NumberFormat.FormatCode = "#,##0";
        chart.AxisY.NumberFormat.IsLinkedToSource = false;

        // Save the document to a DOCX file.
        doc.Save("FormattedAxis.docx");
    }
}
