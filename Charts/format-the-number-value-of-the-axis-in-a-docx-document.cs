using Aspose.Words;
using Aspose.Words.Drawing; // <-- added
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart of size 500x300 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a custom series with categories and numeric values.
        chart.Series.Add(
            "Sample Series",
            new[] { "A", "B", "C", "D" },
            new double[] { 12345, 67890, 23456, 78901 });

        // ------------------------------------------------------------
        // Format the Y‑axis tick labels.
        // ------------------------------------------------------------
        // Access the NumberFormat object of the Y axis and set a custom format code.
        chart.AxisY.NumberFormat.FormatCode = "#,##0";

        // Disable linking to the source cell so the custom format is used.
        chart.AxisY.NumberFormat.IsLinkedToSource = false;

        // Save the document to a DOCX file.
        doc.Save("AxisNumberFormat.docx");
    }
}
