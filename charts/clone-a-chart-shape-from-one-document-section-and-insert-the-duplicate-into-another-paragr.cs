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

        // First paragraph – will contain the original chart.
        builder.Writeln("Original chart paragraph:");

        // Insert a column chart and keep a reference to its shape.
        Shape originalChartShape = builder.InsertChart(ChartType.Column, 432, 252);

        // Verify that the shape indeed contains a chart before proceeding.
        if (!originalChartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Paragraph where the cloned chart will be placed.
        builder.Writeln("Paragraph for the cloned chart:");

        // Deep clone the original chart shape (including its chart data).
        Shape clonedChartShape = (Shape)originalChartShape.Clone(true);

        // Locate the target paragraph (the one we just added).
        Paragraph targetParagraph = (Paragraph)doc.GetChild(
            NodeType.Paragraph,
            doc.GetChildNodes(NodeType.Paragraph, true).Count - 1,
            true);

        // Move the builder to the target paragraph and insert the cloned chart before it.
        builder.MoveTo(targetParagraph);
        builder.InsertNode(clonedChartShape);

        // Save the resulting document.
        doc.Save("ClonedChartExample.docx");
    }
}
