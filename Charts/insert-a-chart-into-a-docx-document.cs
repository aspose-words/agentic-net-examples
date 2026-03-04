using System;
using Aspose.Words;
using Aspose.Words.Drawing; // Added for Shape
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart sized 300x300 pixels (converted to points).
        Shape chartShape = builder.InsertChart(
            ChartType.Pie,
            ConvertUtil.PixelToPoint(300),
            ConvertUtil.PixelToPoint(300));

        // Access the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Remove any demo data that Aspose.Words may have added.
        chart.Series.Clear();

        // Add a series with categories (X values) and corresponding Y values.
        chart.Series.Add(
            "My fruit",
            new[] { "Apples", "Bananas", "Cherries" },
            new[] { 1.3, 2.2, 1.5 });

        // Save the document to a DOCX file.
        doc.Save("ChartDocument.docx");
    }
}
