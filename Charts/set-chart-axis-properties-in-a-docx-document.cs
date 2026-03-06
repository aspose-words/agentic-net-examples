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

        // Insert a column chart of size 500x300 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series that Aspose adds by default.
        chart.Series.Clear();

        // Add a custom series with categories and values.
        chart.Series.Add("Sample Series",
            new[] { "A", "B", "C", "D" },
            new double[] { 120, 80, 150, 60 });

        // ----- Configure the X axis -----
        ChartAxis xAxis = chart.AxisX;
        xAxis.CategoryType = AxisCategoryType.Category;      // Use explicit categories.
        xAxis.Crosses = AxisCrosses.Minimum;                // Axis crosses at the minimum.
        xAxis.ReverseOrder = false;                         // Normal order.
        xAxis.MajorTickMark = AxisTickMark.Inside;          // Inside major ticks.
        xAxis.MinorTickMark = AxisTickMark.Cross;           // Cross minor ticks.
        xAxis.MajorUnit = 10.0;                             // Distance between major ticks.
        xAxis.MinorUnit = 5.0;                              // Distance between minor ticks.
        xAxis.TickLabels.Offset = 30;                       // Distance of labels from axis.
        xAxis.TickLabels.Position = AxisTickLabelPosition.Low;
        xAxis.TickLabels.IsAutoSpacing = false;            // Manual spacing.
        xAxis.TickMarkSpacing = 1;                          // One tick per category.

        // ----- Configure the Y axis -----
        ChartAxis yAxis = chart.AxisY;
        yAxis.CategoryType = AxisCategoryType.Automatic;    // Let Word decide.
        yAxis.Crosses = AxisCrosses.Maximum;                // Crosses at the maximum.
        yAxis.ReverseOrder = true;                          // Values displayed high‑to‑low.
        yAxis.MajorTickMark = AxisTickMark.Inside;
        yAxis.MinorTickMark = AxisTickMark.Cross;
        yAxis.MajorUnit = 50.0;
        yAxis.MinorUnit = 10.0;
        yAxis.TickLabels.Position = AxisTickLabelPosition.NextToAxis;
        yAxis.TickLabels.Alignment = ParagraphAlignment.Center;
        yAxis.TickLabels.Font.Color = Color.Red;            // Red label font.
        yAxis.TickLabels.Spacing = 1;                       // One label per tick.

        // Save the document containing the chart with customized axes.
        doc.Save("ChartAxisProperties.docx");
    }
}
