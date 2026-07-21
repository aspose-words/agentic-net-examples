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

        // Insert a chart with automatic scaling (width and height set to 0 request 100% scale).
        Shape chartShape = builder.InsertChart(ChartType.Column, 0, 0);
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a simple data series.
        chart.Series.Add(
            "Sales",
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new[] { 150.0, 200.0, 250.0, 300.0 });

        // Add a visible title.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // Save the initial document.
        doc.Save("chart-auto-scale.docx");

        // Change the page size of the first section (e.g., to A4) and adjust margins.
        builder.PageSetup.PaperSize = PaperSize.A4;
        builder.PageSetup.TopMargin = ConvertUtil.InchToPoint(1);
        builder.PageSetup.BottomMargin = ConvertUtil.InchToPoint(1);
        builder.PageSetup.LeftMargin = ConvertUtil.InchToPoint(1);
        builder.PageSetup.RightMargin = ConvertUtil.InchToPoint(1);

        // Rebuild the layout so the chart rescales to fit the new page dimensions.
        doc.UpdatePageLayout();

        // Save the document after the page size change.
        doc.Save("chart-auto-scale-updated.docx");
    }
}
