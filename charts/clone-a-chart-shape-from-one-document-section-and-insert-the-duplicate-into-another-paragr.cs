using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the first paragraph.
        Shape originalChart = builder.InsertChart(ChartType.Column, 432, 252);

        // Add a paragraph that will contain the cloned chart.
        builder.Writeln(); // Paragraph break.
        builder.Writeln("Paragraph that will contain the cloned chart:");

        // Clone the original chart shape (deep clone).
        Shape clonedChart = (Shape)originalChart.Clone(true);

        // Move the builder to the paragraph we just added.
        builder.MoveTo(builder.CurrentParagraph);

        // Insert the cloned chart into the current paragraph.
        builder.InsertNode(clonedChart);

        // Save the document.
        doc.Save("ClonedChart.docx");
    }
}
