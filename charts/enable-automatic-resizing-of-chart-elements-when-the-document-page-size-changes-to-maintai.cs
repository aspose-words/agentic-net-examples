using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts; // Required for Chart, ChartType, etc.

public class ChartAutoResizeExample
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with an initial size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Populate the chart with sample data.
        chart.Series.Clear();
        chart.Series.Add("Quarterly Sales", new double[] { 15000, 21000, 18000, 24000 });

        // Save the document before changing the page size (optional step to show original layout).
        doc.Save("chart-original.docx");

        // Change the page size of the first section to A5.
        PageSetup pageSetup = doc.FirstSection.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;

        // Calculate new dimensions for the chart based on the new page size.
        // Here we set the chart width to 80% of the page width and height to 50% of the page height.
        double newChartWidth = pageSetup.PageWidth * 0.8;
        double newChartHeight = pageSetup.PageHeight * 0.5;

        // Apply the new dimensions to the chart shape.
        chartShape.Width = newChartWidth;
        chartShape.Height = newChartHeight;

        // Save the updated document where the chart has been resized automatically.
        doc.Save("chart-resized.docx");
    }
}
