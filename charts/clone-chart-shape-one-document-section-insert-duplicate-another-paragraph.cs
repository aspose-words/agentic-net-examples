using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace CloneChartShapeExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // ------------------------------------------------------------
            // Create a source document and insert a chart shape into it.
            // ------------------------------------------------------------
            Document srcDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

            // Insert a simple pie chart (300x300 points).
            Shape srcChartShape = srcBuilder.InsertChart(
                ChartType.Pie,
                ConvertUtil.PixelToPoint(300),
                ConvertUtil.PixelToPoint(300));

            // Optional: customize the chart (title, series, etc.).
            Chart srcChart = srcChartShape.Chart;
            srcChart.Title.Text = "Source Chart";
            srcChart.Title.Show = true;

            // Save the source document (demonstrates the create‑save rule).
            srcDoc.Save("SourceChart.docx");

            // ------------------------------------------------------------
            // Create a destination document with a couple of paragraphs.
            // ------------------------------------------------------------
            Document dstDoc = new Document();
            DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);

            // First paragraph – will stay unchanged.
            dstBuilder.Writeln("Paragraph before the chart.");

            // Second paragraph – the chart will be inserted here.
            Paragraph targetParagraph = dstBuilder.CurrentParagraph;

            // Third paragraph – after the chart.
            dstBuilder.Writeln("Paragraph after the chart.");

            // ------------------------------------------------------------
            // Clone the chart shape from the source document.
            // ------------------------------------------------------------
            // Find the first shape that contains a chart in the source document.
            Shape chartShapeToClone = (Shape)srcDoc.GetChildNodes(NodeType.Shape, true)
                .Cast<Shape>()
                .First(s => s.HasChart);

            // Perform a deep clone of the shape (including its child nodes).
            Shape clonedShape = (Shape)chartShapeToClone.Clone(true);

            // Import the cloned shape into the destination document.
            // This resolves style, list and other document‑specific references.
            Node importedShapeNode = dstDoc.ImportNode(clonedShape, true);

            // ------------------------------------------------------------
            // Insert the imported chart shape into the target paragraph.
            // ------------------------------------------------------------
            // Move the builder to the paragraph where we want the chart.
            dstBuilder.MoveTo(targetParagraph);

            // Insert the shape node. The shape will become a child of the paragraph.
            dstBuilder.InsertNode(importedShapeNode);

            // Save the resulting document.
            dstDoc.Save("DestinationWithClonedChart.docx");
        }
    }
}
