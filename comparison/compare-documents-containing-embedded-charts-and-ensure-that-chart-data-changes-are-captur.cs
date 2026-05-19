using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create the original document with a chart.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        Shape originalChartShape = builderOriginal.InsertChart(ChartType.Column, 400, 300);
        Chart originalChart = originalChartShape.Chart;
        originalChart.Series.Clear();
        originalChart.Series.Add("Series 1",
            new[] { "Category A", "Category B" },
            new[] { 10.0, 20.0 });

        // Create the revised document with the same chart but different data.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        Shape revisedChartShape = builderRevised.InsertChart(ChartType.Column, 400, 300);
        Chart revisedChart = revisedChartShape.Chart;
        revisedChart.Series.Clear();
        revisedChart.Series.Add("Series 1",
            new[] { "Category A", "Category B" },
            new[] { 30.0, 40.0 }); // Changed data

        // Ensure both documents have no revisions before comparison.
        if (original.HasRevisions || revised.HasRevisions)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the documents. Revisions will be added to the original document.
        original.Compare(revised, "ChartComparer", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparing documents with chart changes.");

        // Count revisions that are related to the chart shape.
        int chartRevisions = original.Revisions.Count(r =>
            r.ParentNode != null && r.ParentNode.NodeType == NodeType.Shape);

        // Output the results.
        Console.WriteLine($"Total revisions detected: {original.Revisions.Count}");
        Console.WriteLine($"Revisions related to chart data changes: {chartRevisions}");

        // Save the compared document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartComparisonResult.docx");
        original.Save(outputPath);
    }
}
