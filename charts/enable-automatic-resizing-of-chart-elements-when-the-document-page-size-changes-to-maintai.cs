using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ChartAutoResizeExample
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with an initial size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 300, 200);
        Chart chart = chartShape.Chart;

        // Add a simple series to the chart.
        chart.Series.Clear();
        chart.Series.Add("Sales", new double[] { 15, 30, 45 });

        // Save the original document.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string originalPath = Path.Combine(outputDir, "original.docx");
        doc.Save(originalPath);

        // Change the page size (e.g., to A4).
        Section section = doc.FirstSection;
        PageSetup pageSetup = section.PageSetup;
        pageSetup.PageWidth = 595.0;   // A4 width in points (210 mm)
        pageSetup.PageHeight = 842.0;  // A4 height in points (297 mm)

        // Automatically resize the chart to occupy 50% of page width and 30% of page height.
        chartShape.Width = pageSetup.PageWidth * 0.5;
        chartShape.Height = pageSetup.PageHeight * 0.3;

        // Center the chart horizontally and vertically.
        chartShape.Left = (pageSetup.PageWidth - chartShape.Width) / 2;
        chartShape.Top = (pageSetup.PageHeight - chartShape.Height) / 2;

        // Save the resized document.
        string resizedPath = Path.Combine(outputDir, "resized.docx");
        doc.Save(resizedPath);
    }
}
