using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class InsertColumnChartExample
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with default data. Width = 432 points, Height = 252 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);

        // The chart already contains default series and categories.
        // No additional data manipulation is required for this simple example.

        // Save the document to the current working directory.
        doc.Save("insert-chart.docx");
    }
}
