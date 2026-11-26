using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with an initial size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        if (!chartShape.HasChart)
            throw new InvalidOperationException("Inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;
        chart.Series.Clear();
        chart.Series.Add("Series 1", new double[] { 10, 20, 30 });

        // Configure initial page size (A4) and margins.
        Section section = doc.FirstSection;
        PageSetup pageSetup = section.PageSetup;
        pageSetup.PageWidth = 595;   // A4 width in points
        pageSetup.PageHeight = 842;  // A4 height in points
        pageSetup.LeftMargin = 50;
        pageSetup.RightMargin = 50;
        pageSetup.TopMargin = 50;
        pageSetup.BottomMargin = 50;

        // Change the page size (e.g., to Letter) to trigger resizing.
        pageSetup.PageWidth = 612;   // Letter width in points
        pageSetup.PageHeight = 792;  // Letter height in points

        // Automatically resize the chart to fit the new page layout.
        ResizeChartToFitPage(chartShape, pageSetup);

        // Save the document.
        doc.Save("ResizedChart.docx");
    }

    private static void ResizeChartToFitPage(Shape chartShape, PageSetup pageSetup)
    {
        if (chartShape == null) throw new ArgumentNullException(nameof(chartShape));
        if (pageSetup == null) throw new ArgumentNullException(nameof(pageSetup));
        if (!chartShape.HasChart) throw new InvalidOperationException("Shape does not contain a chart.");

        // Calculate the usable area inside the margins.
        double usableWidth = pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin;
        double usableHeight = pageSetup.PageHeight - pageSetup.TopMargin - pageSetup.BottomMargin;

        // Apply a proportion of the usable area to the chart dimensions.
        chartShape.Width = usableWidth * 0.8;   // 80% of the available width
        chartShape.Height = usableHeight * 0.5; // 50% of the available height
    }
}
