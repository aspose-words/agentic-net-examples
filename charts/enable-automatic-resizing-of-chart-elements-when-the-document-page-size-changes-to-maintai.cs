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

        // Insert a column chart with an initial size (width: 432 points, height: 252 points).
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Add a simple title to the chart.
        chart.Title.Text = "Sales Chart";
        chart.Title.Show = true;

        // Save the document with the original chart size (optional step to see the before state).
        doc.Save("chart_initial.docx");

        // Change the page size of the document to A4.
        builder.PageSetup.PaperSize = PaperSize.A4;

        // Rebuild the page layout after changing the page setup.
        doc.UpdatePageLayout();

        // Calculate new dimensions for the chart to keep it proportional to the new page size.
        // Here we use 80% of the page width and preserve the original aspect ratio.
        double pageWidth = builder.PageSetup.PageWidth;
        double originalAspectRatio = chartShape.Height / chartShape.Width;
        double newChartWidth = pageWidth * 0.8;
        double newChartHeight = newChartWidth * originalAspectRatio;

        // Ensure the shape actually contains a chart before resizing.
        if (chartShape.HasChart)
        {
            chartShape.Width = newChartWidth;
            chartShape.Height = newChartHeight;
        }

        // Save the document with the resized chart.
        doc.Save("chart_resized.docx");
    }
}
