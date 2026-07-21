using System;
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

        // Insert a chart into the first paragraph of the document.
        Shape originalChart = builder.InsertChart(ChartType.Column, 432, 252);
        // Optionally give the chart a name for identification.
        originalChart.Name = "OriginalChart";

        // Add a new paragraph where the cloned chart will be placed.
        builder.Writeln(); // Creates an empty paragraph and moves the cursor into it.

        // Clone the original chart shape. The Clone method returns a Node, so cast it back to Shape.
        Shape clonedChart = (Shape)originalChart.Clone(true);

        // Insert the cloned chart at the current cursor position (inside the new paragraph).
        builder.InsertNode(clonedChart);

        // Save the document to the local file system.
        doc.Save("ClonedChart.docx");
    }
}
