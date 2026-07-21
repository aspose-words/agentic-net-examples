using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class CompareChartRevisions
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the sample documents.
        string originalPath = Path.Combine(outputDir, "Original.docx");
        string revisedPath = Path.Combine(outputDir, "Revised.docx");
        string comparedPath = Path.Combine(outputDir, "Compared.docx");

        // ---------- Create the original document with an embedded chart ----------
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Document containing a chart.");

        // Insert a column chart and add a data series.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;
        chart.Series.Clear();
        chart.Series.Add("Series 1", new[] { "A", "B", "C" }, new[] { 10.0, 20.0, 30.0 });

        // Save the original document.
        original.Save(originalPath);

        // ---------- Create the revised document by cloning the original ----------
        Document revised = (Document)original.Clone(true);

        // Locate the chart shape in the revised document.
        Shape revisedChartShape = (Shape)revised.GetChild(NodeType.Shape, 0, true);
        Chart revisedChart = revisedChartShape.Chart;

        // Modify the chart data to simulate a change.
        // Instead of accessing a non‑existent Values property, recreate the series with new data.
        revisedChart.Series.Clear();
        revisedChart.Series.Add("Series 1", new[] { "A", "B", "C" }, new[] { 15.0, 25.0, 35.0 });

        // Save the revised document.
        revised.Save(revisedPath);

        // ---------- Compare the original with the revised document ----------
        // The comparison will generate revisions for any differences, including chart data changes.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Save the document that now contains the revisions.
        original.Save(comparedPath);

        // ---------- Inspect revisions ----------
        int totalRevisions = original.Revisions.Count;
        Console.WriteLine($"Total revisions after comparison: {totalRevisions}");

        // Count revisions that are related to the chart shape.
        int chartRevisions = 0;
        foreach (Revision rev in original.Revisions)
        {
            if (rev.ParentNode != null && rev.ParentNode.NodeType == NodeType.Shape)
                chartRevisions++;
        }
        Console.WriteLine($"Revisions related to chart data: {chartRevisions}");
    }
}
