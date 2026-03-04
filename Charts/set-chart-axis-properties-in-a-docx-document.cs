using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;               // <-- added for Shape
using Aspose.Words.Drawing.Charts;

class ChartAxisPropertiesExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart shape.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series that Aspose adds by default.
        chart.Series.Clear();

        // Add a custom series with categories (X‑axis) and values (Y‑axis).
        chart.Series.Add("Sample Series",
            new[] { "Apple", "Banana", "Cherry", "Date", "Elderberry" },
            new double[] { 120, 80, 150, 60, 200 });

        // ----- Configure X‑axis (primary category axis) -----
        ChartAxis xAxis = chart.AxisX;
        xAxis.CategoryType = AxisCategoryType.Category;   // Explicit category axis.
        xAxis.Crosses = AxisCrosses.Minimum;              // Crosses the Y‑axis at its minimum.
        xAxis.ReverseOrder = false;                       // Normal order.
        xAxis.MajorTickMark = AxisTickMark.Inside;        // Inside major tick marks.
        xAxis.MinorTickMark = AxisTickMark.Cross;         // Cross minor tick marks.
        xAxis.MajorUnit = 10.0;                           // Distance between major ticks.
        xAxis.MinorUnit = 5.0;                            // Distance between minor ticks.
        xAxis.TickLabels.Offset = 30;                     // Move labels away from axis.
        xAxis.TickLabels.Position = AxisTickLabelPosition.Low;
        xAxis.TickLabels.IsAutoSpacing = false;           // Manual spacing.
        xAxis.TickMarkSpacing = 1;                        // Tick mark interval.

        // ----- Configure Y‑axis (primary value axis) -----
        ChartAxis yAxis = chart.AxisY;
        yAxis.CategoryType = AxisCategoryType.Automatic; // Let Word decide.
        yAxis.Crosses = AxisCrosses.Maximum;             // Crosses the X‑axis at its maximum.
        yAxis.ReverseOrder = true;                       // Values displayed from max to min.
        yAxis.MajorTickMark = AxisTickMark.Inside;
        yAxis.MinorTickMark = AxisTickMark.Cross;
        yAxis.MajorUnit = 50.0;
        yAxis.MinorUnit = 10.0;
        yAxis.TickLabels.Position = AxisTickLabelPosition.NextToAxis;
        yAxis.TickLabels.Alignment = ParagraphAlignment.Center;
        yAxis.TickLabels.Font.Color = Color.Red;
        yAxis.TickLabels.Spacing = 1;

        // Optional: set axis titles.
        ChartAxisTitle xTitle = chart.AxisX.Title;
        xTitle.Text = "Fruits";
        xTitle.Show = true;

        ChartAxisTitle yTitle = chart.AxisY.Title;
        yTitle.Text = "Quantity";
        yTitle.Show = true;
        yTitle.Font.Size = 12;
        yTitle.Font.Color = Color.Blue;

        // Save the document to disk.
        doc.Save("Charts.AxisProperties.docx");
    }
}
