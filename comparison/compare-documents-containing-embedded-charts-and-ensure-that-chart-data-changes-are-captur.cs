using System;
using System.IO;
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
        builderOriginal.Writeln("Original document with chart:");
        Shape originalChartShape = builderOriginal.InsertChart(ChartType.Column, 400, 300);
        Chart originalChart = originalChartShape.Chart;
        originalChart.Series.Clear();
        originalChart.Series.Add("Series 1", new[] { "A", "B", "C" }, new[] { 1.0, 2.0, 3.0 });

        // Save the original document (optional, for inspection).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        original.Save(originalPath);

        // Create the revised document with a modified chart (change data point for category B).
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Revised document with chart:");
        Shape revisedChartShape = builderRevised.InsertChart(ChartType.Column, 400, 300);
        Chart revisedChart = revisedChartShape.Chart;
        revisedChart.Series.Clear();
        revisedChart.Series.Add("Series 1", new[] { "A", "B", "C" }, new[] { 1.0, 5.0, 3.0 }); // B changed from 2 to 5

        // Save the revised document (optional, for inspection).
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "revised.docx");
        revised.Save(revisedPath);

        // Compare the original document to the revised document.
        original.Compare(revised, "ChartComparer", DateTime.Now);

        // Count revisions detected after comparison.
        int revisionCount = original.Revisions.Count;

        // Save the comparison result.
        string comparedPath = Path.Combine(Directory.GetCurrentDirectory(), "compared.docx");
        original.Save(comparedPath);

        // Output revision information.
        Console.WriteLine($"Revisions detected: {revisionCount}");
        foreach (Revision revision in original.Revisions)
        {
            Console.WriteLine($"- Type: {revision.RevisionType}, Author: {revision.Author}, Date: {revision.DateTime}");
        }
    }
}
