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

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart with a size of 300x300 pixels.
        Shape chartShape = builder.InsertChart(
            ChartType.Pie,
            ConvertUtil.PixelToPoint(300),
            ConvertUtil.PixelToPoint(300));

        // Access the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a series with categories and corresponding values.
        chart.Series.Add(
            "Fruits",
            new[] { "Apples", "Bananas", "Cherries" },
            new[] { 1.3, 2.2, 1.5 });

        // Set a visible title for the chart.
        chart.Title.Text = "Fruit Distribution";
        chart.Title.Show = true;

        // Save the document as a DOCX file.
        doc.Save("ChartDocument.docx");
    }
}
