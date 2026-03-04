using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class ChartDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a specific size.
        // The InsertChart method returns a Shape that contains the Chart object.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Define categories (X‑axis) and corresponding values (Y‑axis).
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 15000, 21000, 18000, 24000 };

        // Add a new series to the chart.
        chart.Series.Add("Revenue", categories, values);

        // Set a title for the chart.
        chart.Title.Text = "Annual Revenue";
        chart.Title.Font.Size = 16;
        chart.Title.Font.Color = Color.DarkBlue;
        chart.Title.Show = true;               // Ensure the title is visible.
        chart.Title.Overlay = false;           // Do not allow other elements to overlap the title.

        // Apply a predefined chart style.
        chart.Style = ChartStyle.ShadedPlot;

        // Format the X axis.
        chart.AxisX.Title.Text = "Quarter";
        chart.AxisX.Title.Font.Size = 12;
        chart.AxisX.Title.Font.Color = Color.Black;
        chart.AxisX.TickLabels.Position = AxisTickLabelPosition.Low;
        chart.AxisX.MajorTickMark = AxisTickMark.Inside;
        chart.AxisX.MinorTickMark = AxisTickMark.Cross;

        // Format the Y axis.
        chart.AxisY.Title.Text = "Revenue (USD)";
        chart.AxisY.Title.Font.Size = 12;
        chart.AxisY.Title.Font.Color = Color.Black;
        chart.AxisY.TickLabels.Position = AxisTickLabelPosition.NextToAxis;
        chart.AxisY.MajorTickMark = AxisTickMark.Inside;
        chart.AxisY.MinorTickMark = AxisTickMark.Cross;

        // Save the document to a DOCX file.
        doc.Save("ChartDemo.docx");
    }
}
