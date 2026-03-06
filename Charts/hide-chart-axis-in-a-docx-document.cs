using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart of size 500x300 points.
        Shape chartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series to start with a clean chart.
        chart.Series.Clear();

        // Add a custom series with categories for the X‑axis and numeric values for the Y‑axis.
        chart.Series.Add("Sample Series",
            new[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
            new[] { 1.2, 0.3, 2.1, 2.9, 4.2 });

        // Hide the primary X and Y axes.
        chart.AxisX.Hidden = true;
        chart.AxisY.Hidden = true;

        // Save the document to a DOCX file.
        doc.Save("HideChartAxis.docx");
    }
}
