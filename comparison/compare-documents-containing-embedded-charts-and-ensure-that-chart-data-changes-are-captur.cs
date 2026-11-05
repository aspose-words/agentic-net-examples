using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ChartComparisonExample
{
    public static void Main()
    {
        // Create the original document with a chart.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        Shape originalChartShape = builderOriginal.InsertChart(ChartType.Column, 400, 300);
        Chart originalChart = originalChartShape.Chart;
        ChartSeries originalSeries = originalChart.Series[0];
        originalSeries.Name = "Original Series";

        original.Save("Original.docx");

        // Create the revised document with modified chart data (change the series name).
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        Shape revisedChartShape = builderRevised.InsertChart(ChartType.Column, 400, 300);
        Chart revisedChart = revisedChartShape.Chart;
        ChartSeries revisedSeries = revisedChart.Series[0];
        revisedSeries.Name = "Revised Series";

        revised.Save("Revised.docx");

        // Compare the documents, capturing revisions.
        original.Compare(revised, "ChartComparer", DateTime.Now);

        // Verify that revisions were created.
        int revisionCount = original.Revisions.Count;
        Console.WriteLine($"Total revisions detected: {revisionCount}");
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Author: {rev.Author}");
        }

        // Save the comparison result.
        original.Save("ChartComparisonResult.docx");
    }
}
