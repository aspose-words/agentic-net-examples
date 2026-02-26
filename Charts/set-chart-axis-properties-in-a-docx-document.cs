using System;
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

        // Remove the default demo series so we can start with a clean chart.
        chart.Series.Clear();

        // Add a custom series with four categories and corresponding values.
        chart.Series.Add("Sample Series",
            new[] { "A", "B", "C", "D" },
            new double[] { 120, 80, 150, 60 });

        // ---------- Configure the X axis ----------
        ChartAxis xAxis = chart.AxisX;
        xAxis.CategoryType = AxisCategoryType.Category;      // Use explicit categories.
        xAxis.Crosses = AxisCrosses.Minimum;                // Cross at the minimum of the Y axis.
        xAxis.ReverseOrder = false;                         // Normal order.
        xAxis.MajorTickMark = AxisTickMark.Inside;          // Inside major tick marks.
        xAxis.MinorTickMark = AxisTickMark.Cross;           // Cross minor tick marks.
        xAxis.MajorUnit = 10.0;                             // Distance between major ticks.
        xAxis.MinorUnit = 5.0;                              // Distance between minor ticks.
        xAxis.TickLabels.Offset = 30;                       // Distance of labels from the axis.
        xAxis.TickLabels.Position = AxisTickLabelPosition.Low;
        xAxis.TickLabels.IsAutoSpacing = false;             // Manual spacing.
        xAxis.TickMarkSpacing = 1;                          // One tick per category.

        // ---------- Configure the Y axis ----------
        ChartAxis yAxis = chart.AxisY;
        yAxis.CategoryType = AxisCategoryType.Automatic;    // Let Word decide the type.
        yAxis.Crosses = AxisCrosses.Maximum;                // Cross at the maximum of the X axis.
        yAxis.ReverseOrder = true;                          // Display values from max to min.
        yAxis.MajorTickMark = AxisTickMark.Inside;
        yAxis.MinorTickMark = AxisTickMark.Cross;
        yAxis.MajorUnit = 50.0;
        yAxis.MinorUnit = 10.0;
        yAxis.TickLabels.Position = AxisTickLabelPosition.NextToAxis;
        yAxis.TickLabels.Alignment = ParagraphAlignment.Center;
        yAxis.TickLabels.Font.Color = Color.Red;            // Red label font.
        yAxis.TickLabels.Spacing = 1;                       // One label per tick.

        // Save the document with the configured chart.
        doc.Save("ChartAxisProperties.docx");
    }
}
