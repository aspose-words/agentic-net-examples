using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart into the first paragraph of the document.
        // The InsertChart method returns the Shape that contains the chart.
        Shape originalChartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart originalChart = originalChartShape.Chart;

        // Optional: give the chart a title so we can see it in the output.
        originalChart.Title.Text = "Original Chart";
        originalChart.Title.Show = true;

        // Add a paragraph after the original chart.
        builder.Writeln();
        builder.Writeln("Below is the cloned chart:");

        // Clone the chart shape. Clone(true) creates a deep copy of the shape node.
        Shape clonedChartShape = (Shape)originalChartShape.Clone(true);

        // Insert the cloned chart shape at the current builder position (inside the new paragraph).
        builder.InsertNode(clonedChartShape);

        // Save the document to the local file system.
        doc.Save("ClonedChartExample.docx");
    }
}
