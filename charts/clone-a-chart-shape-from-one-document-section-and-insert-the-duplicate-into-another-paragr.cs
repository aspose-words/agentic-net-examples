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

        // Add a paragraph and insert the original chart.
        builder.Writeln("Original chart:");
        Shape originalChart = builder.InsertChart(ChartType.Column, 432, 252);
        originalChart.Name = "OriginalChart";

        // Add a paragraph where the cloned chart will be placed.
        builder.Writeln("Cloned chart:");

        // Verify that the shape actually contains a chart before cloning.
        if (!originalChart.HasChart)
            throw new InvalidOperationException("The shape does not contain a chart.");

        // Clone the chart shape (deep clone) and insert it at the current position.
        Shape clonedChart = (Shape)originalChart.Clone(true);
        builder.InsertNode(clonedChart);

        // Save the resulting document.
        doc.Save("ClonedChart.docx");
    }
}
