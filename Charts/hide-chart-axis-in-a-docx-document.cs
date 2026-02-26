using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart into the document.
        Shape shape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = shape.Chart;

        // Clear the default demo series.
        chart.Series.Clear();

        // Add a custom series (categories for X‑axis, values for Y‑axis).
        chart.Series.Add("Sample Series",
            new[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
            new[] { 1.2, 0.3, 2.1, 2.9, 4.2 });

        // Hide both primary axes.
        chart.AxisX.Hidden = true;
        chart.AxisY.Hidden = true;

        // Save the document.
        doc.Save("HideChartAxis.docx");
    }
}
