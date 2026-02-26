using System;
using Aspose.Words;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using System.Drawing;

class AxisLabelSpacingExample
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series and add a custom series.
        chart.Series.Clear();
        chart.Series.Add("Sample Series",
            new[] { "Category A", "Category B", "Category C", "Category D" },
            new double[] { 120, 80, 150, 60 });

        // Access the X axis.
        ChartAxis xAxis = chart.AxisX;

        // Disable automatic spacing and set a custom interval between tick labels.
        xAxis.TickLabels.IsAutoSpacing = false;   // turn off auto‑spacing
        xAxis.TickLabels.Spacing = 2;             // draw a label every 2 categories

        // (Optional) Adjust other axis properties for clarity.
        xAxis.TickLabels.Position = AxisTickLabelPosition.Low;
        xAxis.TickLabels.Alignment = ParagraphAlignment.Center;

        // Save the document.
        doc.Save("AxisLabelSpacing.docx");
    }
}
