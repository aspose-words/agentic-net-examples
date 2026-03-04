using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class ChartExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a specific size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        // Access the Chart object from the shape.
        Chart chart = chartShape.Chart;

        // Remove the demo data series that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Add a new series with categories (X‑axis) and values (Y‑axis).
        chart.Series.Add("Sales 2023",
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new double[] { 15000, 21000, 18000, 24000 });

        // Set the chart title and make it visible.
        ChartTitle title = chart.Title;
        title.Text = "Quarterly Sales 2023";
        title.Font.Size = 14;
        title.Font.Color = Color.DarkBlue;
        title.Show = true;
        title.Overlay = false; // Keep other elements from overlapping the title.

        // Format the chart background.
        chart.Format.Fill.Solid(Color.LightYellow);

        // Customize the X axis.
        ChartAxis xAxis = chart.AxisX;
        xAxis.Title.Show = true;
        xAxis.Title.Text = "Quarter";
        xAxis.Title.Font.Color = Color.DarkGreen;
        xAxis.MajorTickMark = AxisTickMark.Inside;
        xAxis.MinorTickMark = AxisTickMark.Cross;
        xAxis.MajorUnit = 1;
        xAxis.MinorUnit = 0.5;
        xAxis.TickLabels.Position = AxisTickLabelPosition.Low;
        xAxis.TickLabels.Font.Color = Color.Brown;

        // Customize the Y axis.
        ChartAxis yAxis = chart.AxisY;
        yAxis.Title.Show = true;
        yAxis.Title.Text = "Revenue (USD)";
        yAxis.Title.Font.Color = Color.DarkRed;
        yAxis.MajorTickMark = AxisTickMark.Inside;
        yAxis.MinorTickMark = AxisTickMark.Cross;
        yAxis.MajorUnit = 5000;
        yAxis.MinorUnit = 2500;
        yAxis.TickLabels.Position = AxisTickLabelPosition.NextToAxis;
        yAxis.TickLabels.Font.Color = Color.Purple;

        // Save the document to a DOCX file.
        doc.Save("ChartWithCustomAxes.docx");
    }
}
