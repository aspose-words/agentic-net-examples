using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class CompareChartsWithRevisions
{
    public static void Main()
    {
        // Prepare a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // ---------- Create the original document with an embedded chart ----------
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);

        // Insert a column chart.
        Shape chartShapeOriginal = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chartOriginal = chartShapeOriginal.Chart;

        // Define series and data for the original chart.
        chartOriginal.Series.Clear();
        // Use the overload that accepts categories and values.
        chartOriginal.Series.Add("Series 1", new[] { "A", "B", "C" }, new double[] { 10, 20, 30 });

        // Save the original document.
        string originalPath = Path.Combine(artifactsDir, "Original.docx");
        docOriginal.Save(originalPath);

        // ---------- Create the edited document by cloning and modifying the chart ----------
        Document docEdited = (Document)docOriginal.Clone(true);
        Shape chartShapeEdited = (Shape)docEdited.GetChild(NodeType.Shape, 0, true);
        Chart chartEdited = chartShapeEdited.Chart;

        // Change the chart data to simulate an edit.
        chartEdited.Series.Clear();
        chartEdited.Series.Add("Series 1", new[] { "A", "B", "C" }, new double[] { 15, 25, 35 });

        // Save the edited document.
        string editedPath = Path.Combine(artifactsDir, "Edited.docx");
        docEdited.Save(editedPath);

        // ---------- Compare the two documents ----------
        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            docOriginal.Compare(docEdited, "ChartComparer", DateTime.Now);
        }

        // Verify that revisions were created (chart data change should be captured).
        int revisionCount = docOriginal.Revisions.Count;
        Console.WriteLine($"Number of revisions detected: {revisionCount}");

        // Save the comparison result (original document now contains revisions).
        string resultPath = Path.Combine(artifactsDir, "ComparisonResult.docx");
        docOriginal.Save(resultPath);
    }
}
