using System;
using System.IO;
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

        // Insert a column chart into the first section.
        Shape originalChartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart originalChart = originalChartShape.Chart;

        // Optional: clear demo data and add custom series.
        originalChart.Series.Clear();
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        originalChart.Series.Add("Sales", categories, new double[] { 150, 200, 180, 220 });

        // Ensure the shape really contains a chart.
        if (!originalChartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Insert a section break to create a second section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Add a paragraph in the new section where the cloned chart will be placed.
        builder.Writeln("Paragraph before the cloned chart.");

        // Clone the original chart shape (deep clone).
        Node clonedChartNode = originalChartShape.Clone(true);

        // Move the builder to the end of the document (after the paragraph just added).
        builder.MoveToDocumentEnd();

        // Insert the cloned chart shape before the current position.
        builder.InsertNode(clonedChartNode);

        // Add another paragraph after the cloned chart for demonstration.
        builder.Writeln("Paragraph after the cloned chart.");

        // Save the document to the working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ClonedChart.docx");
        doc.Save(outputPath);
    }
}
