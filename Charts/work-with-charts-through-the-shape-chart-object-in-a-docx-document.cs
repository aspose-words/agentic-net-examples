using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace AsposeWordsChartDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to insert a column chart.
            DocumentBuilder builder = new DocumentBuilder(doc);
            Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
            Chart chart = chartShape.Chart;

            // Clear the demo data series that are added by default.
            chart.Series.Clear();

            // Add a new series with categories (X‑axis) and values (Y‑axis).
            chart.Series.Add("Sales",
                new[] { "Q1", "Q2", "Q3", "Q4" },
                new double[] { 15000, 21000, 18000, 24000 });

            // Set the chart title.
            ChartTitle title = chart.Title;
            title.Text = "Quarterly Sales";
            title.Font.Size = 14;
            title.Font.Color = Color.DarkBlue;
            title.Show = true;          // Make sure the title is visible.
            title.Overlay = false;      // Do not allow other elements to overlap the title.

            // Format the chart background.
            chart.Format.Fill.Solid(Color.LightYellow);

            // Hide axis tick labels for a cleaner look.
            chart.AxisX.TickLabels.Position = AxisTickLabelPosition.None;
            chart.AxisY.TickLabels.Position = AxisTickLabelPosition.None;

            // Change the fill color of the data series.
            chart.Series[0].Format.Fill.Solid(Color.CornflowerBlue);

            // Save the document to a DOCX file.
            doc.Save("ChartDemo.docx");
        }
    }
}
