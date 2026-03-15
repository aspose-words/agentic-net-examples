using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart with a width and height of 300 pixels (converted to points).
        Shape chartShape = builder.InsertChart(
            ChartType.Pie,
            ConvertUtil.PixelToPoint(300),
            ConvertUtil.PixelToPoint(300));

        // Retrieve the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Remove the default demo series that Aspose.Words adds.
        chart.Series.Clear();

        // Add a new series with categories and corresponding values.
        chart.Series.Add(
            "Fruits",
            new[] { "Apples", "Bananas", "Cherries" },
            new[] { 1.3, 2.2, 1.5 });

        // Save the document to a DOCX file.
        doc.Save("ChartDocument.docx");
    }
}
