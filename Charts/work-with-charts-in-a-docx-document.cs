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

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a width of 500 points and a height of 300 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the automatically generated demo series.
        chart.Series.Clear();

        // Define categories (X‑axis) and corresponding values (Y‑axis).
        string[] categories = { "Word", "PDF", "Excel", "GoogleDocs", "Note" };
        double[] values = { 640, 320, 280, 120, 150 };

        // Add a new series to the chart.
        chart.Series.Add("Aspose Test Series", categories, values);

        // Set and format the chart title.
        chart.Title.Text = "Document Conversion Statistics";
        chart.Title.Font.Size = 14;
        chart.Title.Font.Color = Color.DarkBlue;
        chart.Title.Show = true; // Ensure the title is visible.

        // Apply a solid fill to the chart background.
        chart.Format.Fill.Solid(Color.LightGray);

        // Hide tick labels on both axes for a cleaner look.
        chart.AxisX.TickLabels.Position = AxisTickLabelPosition.None;
        chart.AxisY.TickLabels.Position = AxisTickLabelPosition.None;

        // Save the document to a DOCX file.
        doc.Save("ChartExample.docx");
    }
}
