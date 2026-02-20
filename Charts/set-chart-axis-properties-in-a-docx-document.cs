using System;
using Aspose.Words;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Drawing;
using System.Drawing;

class ChartAxisDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart of size 500x300 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series that Aspose adds by default.
        chart.Series.Clear();

        // Add a custom series with categories (X‑axis) and values (Y‑axis).
        chart.Series.Add("Sample Series",
            new[] { "A", "B", "C", "D", "E" },
            new double[] { 120, 80, 150, 200, 90 });

        // ----- Configure X axis -----
        ChartAxis xAxis = chart.AxisX;
        xAxis.CategoryType = AxisCategoryType.Category;   // ordinary categories
        xAxis.Crosses = AxisCrosses.Minimum;              // cross at the minimum of Y axis
        xAxis.ReverseOrder = false;                       // normal order
        xAxis.MajorTickMark = AxisTickMark.Inside;        // major ticks inside plot area
        xAxis.MinorTickMark = AxisTickMark.Cross;         // minor ticks cross the axis
        xAxis.MajorUnit = 1;                              // one major tick per category
        xAxis.MinorUnit = 0.5;                            // half‑category minor ticks
        xAxis.TickLabels.Offset = 30;                     // distance from axis
        xAxis.TickLabels.Position = AxisTickLabelPosition.Low;
        xAxis.TickLabels.IsAutoSpacing = false;
        xAxis.TickMarkSpacing = 1;                        // draw a label for each category

        // ----- Configure Y axis -----
        ChartAxis yAxis = chart.AxisY;
        yAxis.CategoryType = AxisCategoryType.Automatic; // value axis
        yAxis.Crosses = AxisCrosses.Maximum;             // cross at the maximum of X axis
        yAxis.ReverseOrder = true;                       // display values from max to min
        yAxis.MajorTickMark = AxisTickMark.Inside;
        yAxis.MinorTickMark = AxisTickMark.Cross;
        yAxis.MajorUnit = 50;                             // major tick every 50 units
        yAxis.MinorUnit = 10;                             // minor tick every 10 units
        yAxis.TickLabels.Position = AxisTickLabelPosition.NextToAxis;
        yAxis.TickLabels.Alignment = ParagraphAlignment.Center;
        yAxis.TickLabels.Font.Color = Color.Red;
        yAxis.TickLabels.Spacing = 1;

        // ----- Add axis titles -----
        ChartAxisTitle xTitle = chart.AxisX.Title;
        xTitle.Text = "Categories";
        xTitle.Show = true;

        ChartAxisTitle yTitle = chart.AxisY.Title;
        yTitle.Text = "Values";
        yTitle.Show = true;
        yTitle.Overlay = true;            // allow other elements to overlap the title
        yTitle.Font.Size = 12;
        yTitle.Font.Color = Color.Blue;

        // Save the document.
        doc.Save("ChartAxisProperties.docx");
    }
}
