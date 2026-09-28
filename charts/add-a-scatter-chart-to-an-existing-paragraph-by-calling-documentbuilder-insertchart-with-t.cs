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

        // Add an initial paragraph.
        builder.Writeln("This paragraph will contain a scatter chart.");

        // Locate the first paragraph in the document.
        Paragraph firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
        if (firstParagraph == null)
            throw new InvalidOperationException("Document does not contain any paragraphs.");

        // Move the builder to the start of the existing paragraph.
        builder.MoveTo(firstParagraph);

        // Insert a scatter chart into the paragraph.
        Shape chartShape = builder.InsertChart(ChartType.Scatter, 432, 252);
        Chart chart = chartShape.Chart;

        // Clear any default series and add custom X and Y values.
        chart.Series.Clear();
        chart.Series.Add("Sample Series", new double[] { 1, 2, 3 }, new double[] { 4, 5, 6 });

        // Save the document.
        doc.Save("ScatterChart.docx");
    }
}
