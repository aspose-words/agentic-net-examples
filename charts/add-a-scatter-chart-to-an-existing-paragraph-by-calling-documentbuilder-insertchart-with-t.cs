using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;   // Required for Chart, ChartType, etc.

public class ScatterChartExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the scatter chart.
        builder.Writeln("This paragraph will contain a scatter chart:");

        // Move the builder back to the paragraph we just created.
        Paragraph paragraph = builder.CurrentParagraph;
        builder.MoveTo(paragraph);

        // Insert a scatter chart into the current paragraph.
        Shape chartShape = builder.InsertChart(ChartType.Scatter, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove any default series that were added automatically.
        chart.Series.Clear();

        // Add a series with X and Y values for the scatter chart.
        double[] xValues = { 1, 2, 3, 4, 5 };
        double[] yValues = { 2, 4, 1, 3, 5 };
        chart.Series.Add("Sample Series", xValues, yValues);

        // Save the document to the local file system.
        doc.Save("ScatterChart.docx");
    }
}
