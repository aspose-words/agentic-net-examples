using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

class FormatChartAxisNumber
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        // Width = 500 points, Height = 300 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a custom series with categories (X‑axis) and large numeric values (Y‑axis).
        chart.Series.Add(
            "Sales",
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new double[] { 1_900_000, 850_000, 2_100_000, 1_500_000 });

        // ------------------------------------------------------------
        // Format the Y‑axis numbers.
        // The NumberFormat property returns a ChartNumberFormat object.
        // Set its FormatCode to a custom pattern (e.g., "#,##0") to control
        // how tick labels are displayed and disable linking to the source cell.
        // ------------------------------------------------------------
        chart.AxisY.NumberFormat.FormatCode = "#,##0";
        chart.AxisY.NumberFormat.IsLinkedToSource = false;

        // (Optional) Format the X‑axis numbers in a similar way.
        // chart.AxisX.NumberFormat.FormatCode = "#,##0";
        // chart.AxisX.NumberFormat.IsLinkedToSource = false;

        // Save the document to a DOCX file.
        doc.Save("FormattedChartAxis.docx");
    }
}
