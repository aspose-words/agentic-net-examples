using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Column chart with a specific size.
        // Width and height are specified in points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Define categories for the X‑axis.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };

        // Add two data series to the chart.
        chart.Series.Add("Product A", categories, new double[] { 120, 150, 170, 200 });
        chart.Series.Add("Product B", categories, new double[] { 80, 130, 160, 190 });

        // Set the chart title.
        ChartTitle title = chart.Title;
        title.Text = "Quarterly Sales Comparison";
        title.Font.Size = 14;
        title.Font.Color = Color.DarkBlue;
        title.Show = true;          // Ensure the title is visible.
        title.Overlay = false;      // Do not allow other elements to overlap the title.

        // Format the chart background.
        chart.Format.Fill.Solid(Color.LightYellow);

        // Hide axis tick labels for a cleaner look.
        chart.AxisX.TickLabels.Position = AxisTickLabelPosition.None;
        chart.AxisY.TickLabels.Position = AxisTickLabelPosition.None;

        // Optionally, format the legend background.
        chart.Legend.Format.Fill.Solid(Color.LightGray);

        // Save the document to a DOCX file.
        doc.Save("ChartExample.docx");
    }
}
