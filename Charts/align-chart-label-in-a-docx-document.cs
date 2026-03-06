using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series and add custom data.
        chart.Series.Clear();
        chart.Series.Add(
            "Sample Series",
            new[] { "Category A", "Category B", "Category C", "Category D" },
            new double[] { 10, 20, 30, 40 });

        // Access the Y‑axis and align its tick‑label text to the right.
        ChartAxis yAxis = chart.AxisY;
        yAxis.TickLabels.Alignment = ParagraphAlignment.Right; // Align axis tick labels.

        // Optional: change the font color to highlight the alignment.
        yAxis.TickLabels.Font.Color = Color.Blue;

        // Save the resulting document.
        doc.Save("AlignedChartLabels.docx");
    }
}
