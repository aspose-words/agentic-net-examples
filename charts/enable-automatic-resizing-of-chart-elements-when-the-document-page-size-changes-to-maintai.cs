using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart with width and height set to 0.
        // Zero values tell Aspose.Words to scale the chart to 100% of its container.
        Shape chartShape = builder.InsertChart(ChartType.Column, 0, 0);
        Chart chart = chartShape.Chart;

        // Add a simple series to the chart.
        chart.Series.Clear();
        chart.Series.Add("Sales", new[] { "Q1", "Q2", "Q3", "Q4" }, new[] { 150.0, 200.0, 180.0, 220.0 });

        // Set a title so we can see the chart in the output.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // Position the chart relative to the page margins.
        chartShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        chartShape.RelativeVerticalPosition   = RelativeVerticalPosition.Margin;
        chartShape.Left   = 0;   // Align to the left margin.
        chartShape.Top    = 0;   // Align to the top margin.
        chartShape.WrapType = WrapType.Inline; // Keep it inline for automatic layout.

        // Save the document with the original page size (Letter).
        doc.Save("chart-original.docx");

        // Change the page size – for example, to A5.
        builder.PageSetup.PaperSize = PaperSize.A5;
        // Optionally change orientation or margins here.
        builder.PageSetup.Orientation = Orientation.Portrait;

        // Rebuild the layout so that the chart can be re‑flowed according to the new page size.
        doc.UpdatePageLayout();

        // Save the document after the page size change.
        doc.Save("chart-auto-resized.docx");
    }
}
