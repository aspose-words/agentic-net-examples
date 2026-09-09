using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonOutput");
        Directory.CreateDirectory(outputDir);

        // Create the original DOC file with a simple 2‑cell table.
        Document originalDoc = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(originalDoc);
        builderOriginal.StartTable();
        builderOriginal.InsertCell();
        builderOriginal.Write("Original Cell 1");
        builderOriginal.InsertCell();
        builderOriginal.Write("Original Cell 2");
        builderOriginal.EndTable();
        string originalPath = Path.Combine(outputDir, "original.doc");
        originalDoc.Save(originalPath);

        // Create the revised DOCX file with a modified table (different text and an extra cell).
        Document revisedDoc = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revisedDoc);
        builderRevised.StartTable();
        builderRevised.InsertCell();
        builderRevised.Write("Edited Cell 1");          // changed text
        builderRevised.InsertCell();
        builderRevised.Write("Original Cell 2");        // unchanged text
        builderRevised.InsertCell();
        builderRevised.Write("New Cell 3");             // extra cell
        builderRevised.EndTable();
        string revisedPath = Path.Combine(outputDir, "revised.docx");
        revisedDoc.Save(revisedPath);

        // Perform comparison. The original document will receive revisions.
        originalDoc.Compare(revisedDoc, "Comparer", DateTime.Now);

        // Verify that at least one revision exists.
        if (originalDoc.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Count revisions that affect any part of a table.
        int tableRevisionCount = originalDoc.Revisions
            .Count(r => r.ParentNode?.GetAncestor(NodeType.Table) != null);

        if (tableRevisionCount == 0)
            throw new InvalidOperationException("Expected at least one table revision, but none were found.");

        // Save the comparison result.
        string resultPath = Path.Combine(outputDir, "comparisonResult.docx");
        originalDoc.Save(resultPath);

        // Output a simple summary.
        Console.WriteLine($"Total revisions detected: {originalDoc.Revisions.Count}");
        Console.WriteLine($"Table revisions detected: {tableRevisionCount}");
        Console.WriteLine($"Comparison result saved to: {resultPath}");
    }
}
