using System.Drawing;
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

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series and add custom data.
        chart.Series.Clear();
        chart.Series.Add("Sample Series",
            new[] { "A", "B", "C", "D" },
            new double[] { 10, 20, 30, 40 });

        // Align the Y‑axis tick labels to the right.
        ChartAxis yAxis = chart.AxisY;
        yAxis.TickLabels.Alignment = ParagraphAlignment.Right;

        // Optional: change the label font color to highlight the alignment.
        yAxis.TickLabels.Font.Color = Color.Blue;

        // Save the document.
        doc.Save("AlignedChartLabels.docx");
    }
}
