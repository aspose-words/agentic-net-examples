using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

class InsertChartExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart with a width and height of 300 pixels each.
        // ConvertUtil.PixelToPoint converts pixel dimensions to points (the unit used by Word).
        Shape chartShape = builder.InsertChart(
            ChartType.Pie,
            ConvertUtil.PixelToPoint(300),
            ConvertUtil.PixelToPoint(300));

        // Get the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Remove any default demo series.
        chart.Series.Clear();

        // Add a new series with categories and corresponding values.
        chart.Series.Add(
            "My fruit",                                   // Series name
            new[] { "Apples", "Bananas", "Cherries" },   // X‑axis categories
            new[] { 1.3, 2.2, 1.5 });                    // Y‑axis values

        // Optionally, set a title for the chart.
        chart.Title.Text = "Fruit Distribution";
        chart.Title.Show = true;

        // Save the document to a DOCX file.
        string outputPath = "ChartDocument.docx";
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
