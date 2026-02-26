using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart shape with the specified size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        // Access the Chart object from the shape.
        Chart chart = chartShape.Chart;

        // Configure the chart title.
        ChartTitle title = chart.Title;
        title.Text = "Sales Report 2023";
        title.Font.Size = 16;
        title.Font.Color = Color.DarkBlue;
        title.Show = true;          // Make sure the title is visible.
        title.Overlay = false;      // Do not allow other elements to overlap the title.

        // Remove the default demo series that Aspose.Words inserts.
        chart.Series.Clear();

        // Add a first data series with categories and values.
        chart.Series.Add(
            "Q1",
            new[] { "Jan", "Feb", "Mar" },
            new double[] { 12000, 15000, 13000 });

        // Add a second data series.
        chart.Series.Add(
            "Q2",
            new[] { "Apr", "May", "Jun" },
            new double[] { 14000, 16000, 15500 });

        // Apply a solid fill to the chart background.
        chart.Format.Fill.Solid(Color.LightYellow);

        // Hide tick labels on both axes for a cleaner look.
        chart.AxisX.TickLabels.Position = AxisTickLabelPosition.None;
        chart.AxisY.TickLabels.Position = AxisTickLabelPosition.None;

        // Save the document containing the chart.
        doc.Save("ChartExample.docx");
    }
}
