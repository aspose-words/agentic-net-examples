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

        // Write a paragraph that will later contain the chart.
        builder.Writeln("Paragraph that will hold a scatter chart.");

        // Retrieve the paragraph we just added.
        Paragraph targetParagraph = doc.FirstSection.Body.FirstParagraph;

        // Move the builder's cursor to the target paragraph.
        builder.MoveTo(targetParagraph);

        // Insert a scatter chart into the paragraph using the InsertChart overload.
        Shape chartShape = builder.InsertChart(ChartType.Scatter, 400, 300);
        Chart chart = chartShape.Chart;

        // Clear the default demo series.
        chart.Series.Clear();

        // Add a custom series with X and Y values.
        chart.Series.Add("Sample Series",
            new double[] { 1.0, 2.0, 3.0, 4.0 },
            new double[] { 10.0, 20.0, 15.0, 30.0 });

        // Save the document to the working directory.
        doc.Save("ScatterChart.docx");
    }
}
