using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class ChartAutoResizeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart. The initial width/height are placeholders;
        // we will bind the shape to the page margins so it resizes automatically.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // Bind the chart size to the page margins.
        chartShape.RelativeHorizontalSize = RelativeHorizontalSize.Margin; // width relative to margins
        chartShape.RelativeVerticalSize   = RelativeVerticalSize.Margin;   // height relative to margins

        // Use 80 % of the available margin width and height.
        chartShape.WidthRelative  = 80; // percent of margin width
        chartShape.HeightRelative = 80; // percent of margin height

        // Optional: add some data to the chart.
        Chart chart = chartShape.Chart;
        chart.Series.Clear();
        // The Series.Add method expects a double[] for the values, not int[].
        chart.Series.Add("Sample", new[] { "Jan", "Feb", "Mar" }, new double[] { 15, 30, 20 });

        // Save the document with the default page size (Letter).
        doc.Save("Chart_Letter.docx");

        // Change the page size – the chart will automatically adapt because of the relative bindings.
        builder.PageSetup.PaperSize = PaperSize.A5; // smaller page size

        // Rebuild the layout so the new page dimensions are taken into account.
        doc.UpdatePageLayout();

        // Save the document after the page size change.
        doc.Save("Chart_A5.docx");
    }
}
