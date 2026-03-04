using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Column chart of size 500x300 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that comes with a new chart.
        chart.Series.Clear();

        // Define categories (X‑axis) and values (Y‑axis) for the series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120.5, 150.0, 130.75, 170.2 };

        // Add a single series to the chart using the appropriate overload.
        ChartSeries series = chart.Series.Add("Fiscal 2025", categories, values);

        // Set the fill colour of the series to a custom colour.
        series.Format.Fill.ForeColor = Color.CornflowerBlue;

        // Enable data labels for the series and show the value for each point.
        series.HasDataLabels = true;
        foreach (ChartDataLabel label in series.DataLabels)
        {
            label.ShowValue = true;
            label.NumberFormat.FormatCode = "0.00";
        }

        // Optionally, change the chart title.
        chart.Title.Text = "Quarterly Revenue";
        chart.Title.Show = true;

        // Save the document to disk.
        doc.Save("ChartSeriesExample.docx");
    }
}
