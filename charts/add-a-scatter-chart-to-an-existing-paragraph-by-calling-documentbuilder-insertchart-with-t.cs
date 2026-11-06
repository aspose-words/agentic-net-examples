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

        // Write a paragraph that will contain the scatter chart.
        builder.Writeln("Paragraph that will contain a scatter chart:");

        // Move the builder to the start of the first paragraph.
        builder.MoveToParagraph(0, 0);

        // Insert a scatter chart into the current paragraph.
        Shape chartShape = builder.InsertChart(ChartType.Scatter, 432, 252);
        Chart chart = chartShape.Chart;

        // Clear any default series and add custom data for the scatter chart.
        chart.Series.Clear();
        chart.Series.Add("Series 1", new double[] { 1, 2, 3 }, new double[] { 4, 5, 6 });

        // Save the document.
        doc.Save("ScatterChart.docx");
    }
}
