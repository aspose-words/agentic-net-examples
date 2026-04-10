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

        // Insert a chart with zero width and height.
        // Zero values request 100 % scaling, so the chart will resize automatically with the page.
        Shape chartShape = builder.InsertChart(ChartType.Column, 0, 0);

        // Verify that the inserted shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Add custom series data.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 1500, 2000, 1800, 2200 };
        chart.Series.Add("Sales", categories, values);

        // Set a visible title for the chart.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // Save the document with the default page size (Letter).
        doc.Save("Chart_DefaultPageSize.docx");

        // Change the page size – for example, to A5 dimensions.
        Section firstSection = doc.FirstSection;
        PageSetup pageSetup = firstSection.PageSetup;
        pageSetup.PageWidth = ConvertUtil.InchToPoint(5.8);   // Approx. 5.8 in (A5 width)
        pageSetup.PageHeight = ConvertUtil.InchToPoint(8.3);  // Approx. 8.3 in (A5 height)

        // Save the document after the page size change.
        // The chart will have automatically resized to fit the new page dimensions.
        doc.Save("Chart_ResizedPage.docx");
    }
}
