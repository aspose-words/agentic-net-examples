using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart of size 500x300 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a custom data series with categories and large numeric values.
        chart.Series.Add(
            "Sample Series",
            new[] { "A", "B", "C", "D" },
            new double[] { 1234567, 2345678, 3456789, 4567890 });

        // Format the Y‑axis (value axis) tick labels.
        // Set a custom number format (e.g., "#,##0") and detach it from the source cell.
        chart.AxisY.NumberFormat.FormatCode = "#,##0";
        chart.AxisY.NumberFormat.IsLinkedToSource = false;

        // Save the document to a DOCX file.
        doc.Save("AxisNumberFormat.docx");
    }
}
