using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart that will automatically scale to the shape size.
        // Width and height set to 0 request 100 % scaling (fills the container).
        Shape chartShape = builder.InsertChart(ChartType.Column, 0, 0);
        // Ensure the chart shape is positioned relative to the page margins.
        chartShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        chartShape.RelativeVerticalPosition   = RelativeVerticalPosition.Margin;
        chartShape.Left   = 0;   // align to left margin
        chartShape.Top    = 0;   // align to top margin
        chartShape.Width  = 0;   // 100 % of the container width
        chartShape.Height = 0;   // 100 % of the container height

        Chart chart = chartShape.Chart;

        // Add a simple series to make the chart visible.
        chart.Series.Clear();
        chart.Series.Add("Sales",
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new[] { 150.0, 200.0, 180.0, 220.0 });

        // Set a title for clarity.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // Change the page size after the chart has been inserted.
        // The chart will resize automatically because its dimensions are set to 100 % of the page width/height.
        builder.PageSetup.PaperSize   = PaperSize.A3;      // Larger page size.
        builder.PageSetup.Orientation = Orientation.Portrait;

        // Rebuild the layout so that the new page size is taken into account.
        doc.UpdatePageLayout();

        // Save the document.
        doc.Save("ChartAutoResize.docx");
    }
}
