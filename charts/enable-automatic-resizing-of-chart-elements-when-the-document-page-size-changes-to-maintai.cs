using System;
using System.IO;
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

        // Insert a column chart with an initial size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Add a simple series to the chart.
        chart.Series.Clear();
        chart.Series.Add("Sales", new[] { "Q1", "Q2", "Q3", "Q4" }, new[] { 150.0, 200.0, 180.0, 220.0 });

        // Set a title for clarity.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // Save the document before changing the page size (optional, for comparison).
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string initialPath = Path.Combine(outputDir, "Chart_Initial.docx");
        doc.Save(initialPath);

        // Change the page size of the current section (e.g., to A5).
        builder.PageSetup.PaperSize = PaperSize.A5;

        // Recalculate the usable page width and height (excluding margins).
        double usableWidth = builder.PageSetup.PageWidth - builder.PageSetup.LeftMargin - builder.PageSetup.RightMargin;
        double usableHeight = builder.PageSetup.PageHeight - builder.PageSetup.TopMargin - builder.PageSetup.BottomMargin;

        // Resize the chart to fit within the new page dimensions while preserving a margin.
        // Here we leave a 20‑point margin around the chart.
        const double chartMargin = 20.0;
        chartShape.Width = Math.Max(0, usableWidth - 2 * chartMargin);
        chartShape.Height = Math.Max(0, usableHeight - 2 * chartMargin);

        // Optionally reposition the chart to be centered on the page.
        chartShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        chartShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        chartShape.Left = (builder.PageSetup.PageWidth - chartShape.Width) / 2;
        chartShape.Top = (builder.PageSetup.PageHeight - chartShape.Height) / 2;

        // Update layout to ensure the changes are reflected.
        doc.UpdatePageLayout();

        // Save the final document with the resized chart.
        string finalPath = Path.Combine(outputDir, "Chart_Resized.docx");
        doc.Save(finalPath);
    }
}
