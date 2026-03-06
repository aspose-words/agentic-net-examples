using System;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart with a size of 300x300 pixels (converted to points).
        Shape chartShape = builder.InsertChart(
            ChartType.Pie,
            ConvertUtil.PixelToPoint(300),
            ConvertUtil.PixelToPoint(300));

        // Get the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Remove any demo data that comes with the chart.
        chart.Series.Clear();

        // Add a new series with categories and corresponding values.
        chart.Series.Add(
            "My fruit",
            new[] { "Apples", "Bananas", "Cherries" },
            new[] { 1.3, 2.2, 1.5 });

        // Save the document to a DOCX file.
        doc.Save("ChartDocument.docx", SaveFormat.Docx);
    }
}
