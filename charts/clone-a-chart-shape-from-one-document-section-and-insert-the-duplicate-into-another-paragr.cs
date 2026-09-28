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

        // Insert a column chart into the first section.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Write a paragraph that will hold the cloned chart.
        builder.Writeln("Cloned chart inserted below:");

        // Clone the original chart shape.
        Shape clonedChartShape = (Shape)chartShape.Clone(true);

        // Insert the cloned chart into the current paragraph.
        Paragraph currentParagraph = builder.CurrentParagraph 
            ?? throw new InvalidOperationException("Current paragraph is null.");
        currentParagraph.AppendChild(clonedChartShape);

        // Save the document.
        doc.Save("ClonedChart.docx");
    }
}
