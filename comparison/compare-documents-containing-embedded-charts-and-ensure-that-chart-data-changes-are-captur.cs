using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class CompareChartsWithRevisions
{
    public static void Main()
    {
        // Create the original document with a chart.
        Document originalDoc = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(originalDoc);
        Shape originalChartShape = builderOriginal.InsertChart(ChartType.Column, 400, 300);
        PopulateChart(originalChartShape, "Original Chart");

        // Create the revised document with the same chart type but a different title.
        Document revisedDoc = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revisedDoc);
        Shape revisedChartShape = builderRevised.InsertChart(ChartType.Column, 400, 300);
        PopulateChart(revisedChartShape, "Revised Chart");

        // Perform comparison. Revisions will be stored in the original document.
        originalDoc.Compare(revisedDoc, "ChartComparer", DateTime.Now);

        // Verify that at least one revision exists (chart title change should be captured).
        int revisionCount = originalDoc.Revisions.Count;
        if (revisionCount == 0)
        {
            throw new InvalidOperationException(
                "Expected revisions due to chart title changes, but none were found.");
        }

        // Save the compared document with revisions.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartComparisonResult.docx");
        originalDoc.Save(outputPath);

        // Output revision information to the console.
        Console.WriteLine($"Comparison completed. Revisions detected: {revisionCount}");
        foreach (Revision rev in originalDoc.Revisions)
        {
            // Use the available DateTime property for the revision timestamp.
            Console.WriteLine($"- Revision Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }
    }

    // Helper method to set a chart title. The title difference is sufficient for the comparison test.
    private static void PopulateChart(Shape chartShape, string title)
    {
        if (chartShape == null) throw new ArgumentNullException(nameof(chartShape));
        if (title == null) throw new ArgumentNullException(nameof(title));

        Chart chart = chartShape.Chart;

        // Set chart title to create a detectable difference.
        chart.Title.Text = title;

        // No need to manipulate chart data for this example; title change alone generates a revision.
    }
}
