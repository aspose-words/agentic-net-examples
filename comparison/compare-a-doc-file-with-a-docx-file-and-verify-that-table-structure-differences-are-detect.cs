using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a working directory for the sample files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonDemo");
        Directory.CreateDirectory(workDir);

        // ---------- Create the original DOC file ----------
        var originalDoc = new Document();
        var builder = new DocumentBuilder(originalDoc);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();
        string originalPath = Path.Combine(workDir, "original.doc");
        originalDoc.Save(originalPath, SaveFormat.Doc);

        // ---------- Create the revised DOCX file with a different table structure ----------
        var revisedDoc = new Document();
        var builder2 = new DocumentBuilder(revisedDoc);
        // First table (same as original)
        builder2.StartTable();
        builder2.InsertCell();
        builder2.Write("Cell 1");
        builder2.InsertCell();
        builder2.Write("Cell 2");
        builder2.EndTable();
        // Additional table to create a structural difference
        builder2.StartTable();
        builder2.InsertCell();
        builder2.Write("Cell 3");
        builder2.InsertCell();
        builder2.Write("Cell 4");
        builder2.EndTable();
        string revisedPath = Path.Combine(workDir, "revised.docx");
        revisedDoc.Save(revisedPath, SaveFormat.Docx);

        // ---------- Load the documents ----------
        var docOriginal = new Document(originalPath);
        var docRevised = new Document(revisedPath);

        // ---------- Perform comparison ----------
        docOriginal.Compare(docRevised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (docOriginal.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison, but none were found.");

        // Count revisions that are related to tables.
        int tableRevisionCount = 0;
        foreach (Revision rev in docOriginal.Revisions)
        {
            if (rev.ParentNode != null && rev.ParentNode.NodeType == NodeType.Table)
                tableRevisionCount++;
        }

        // Output the results.
        Console.WriteLine($"Total revisions detected: {docOriginal.Revisions.Count}");
        Console.WriteLine($"Table-related revisions detected: {tableRevisionCount}");

        // Save the comparison result.
        string resultPath = Path.Combine(workDir, "comparisonResult.docx");
        docOriginal.Save(resultPath, SaveFormat.Docx);
    }
}
