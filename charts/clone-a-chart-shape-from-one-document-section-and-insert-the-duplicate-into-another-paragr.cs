using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace ChartCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph and a chart in the first section.
            builder.Writeln("Section 1 - Original chart:");
            Shape originalChartShape = builder.InsertChart(ChartType.Column, 400, 300);
            Chart originalChart = originalChartShape.Chart;
            originalChart.Title.Text = "Original Chart";
            originalChart.Title.Show = true;

            // Add a new section for the target paragraph.
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Section 2 - Target paragraph for cloned chart:");

            // Clone the chart shape (deep clone to copy all child nodes).
            Shape clonedChartShape = (Shape)originalChartShape.Clone(true);

            // Move the builder to the end of the document (after the paragraph we just wrote).
            builder.MoveToDocumentEnd();

            // Insert the cloned chart shape at the current position.
            builder.InsertNode(clonedChartShape);

            // Save the resulting document.
            doc.Save("ChartCloneExample.docx");
        }
    }
}
