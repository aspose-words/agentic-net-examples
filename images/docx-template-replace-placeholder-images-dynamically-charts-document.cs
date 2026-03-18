using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts; // <-- added namespace for ChartType

class Program
{
    static void Main()
    {
        // Load the DOCX template from disk.
        Document doc = new Document("Template.docx");

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart at the bookmark named "Chart1".
        InsertChartAtBookmark(builder, "Chart1", ChartType.Column, 400, 300);

        // Insert a pie chart at the bookmark named "Chart2".
        InsertChartAtBookmark(builder, "Chart2", ChartType.Pie, 300, 300);

        // Save the modified document.
        doc.Save("Result.docx");
    }

    // Helper method that moves to a bookmark and inserts a chart.
    static void InsertChartAtBookmark(DocumentBuilder builder, string bookmarkName, ChartType chartType, double width, double height)
    {
        // Verify that the bookmark exists.
        if (builder.Document.Range.Bookmarks[bookmarkName] != null)
        {
            // Position the cursor at the bookmark.
            builder.MoveToBookmark(bookmarkName);

            // Insert the chart with the specified size.
            Shape chartShape = builder.InsertChart(chartType, width, height);

            // Populate the chart with a simple data series.
            var chart = chartShape.Chart;
            chart.Series.Clear();
            chart.Series.Add("Series 1",
                new[] { "Category A", "Category B", "Category C" },
                new[] { 10.0, 20.0, 30.0 });
        }
    }
}
