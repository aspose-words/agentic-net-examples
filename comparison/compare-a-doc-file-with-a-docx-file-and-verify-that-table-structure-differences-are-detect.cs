using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        const string originalDocPath = "Original.doc";
        const string modifiedDocxPath = "Modified.docx";
        const string comparisonResultPath = "ComparisonResult.docx";

        // -----------------------------------------------------------------
        // Create the original DOC file with a simple 2‑cell table.
        // -----------------------------------------------------------------
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);

        builder.StartTable();
        builder.InsertCell();
        builder.Write("Original Cell 1");
        builder.InsertCell();
        builder.Write("Original Cell 2");
        builder.EndTable();

        // Save as .doc (legacy format).
        originalDoc.Save(originalDocPath);

        // -----------------------------------------------------------------
        // Create the modified DOCX file with a different table layout.
        // Here we change the text of the first cell and add a third cell.
        // -----------------------------------------------------------------
        Document modifiedDoc = new Document();
        builder = new DocumentBuilder(modifiedDoc);

        builder.StartTable();
        builder.InsertCell();
        builder.Write("Modified Cell 1"); // Text changed.
        builder.InsertCell();
        builder.Write("Original Cell 2"); // Same as original.
        builder.InsertCell();
        builder.Write("New Cell 3"); // Additional cell.
        builder.EndTable();

        // Save as .docx.
        modifiedDoc.Save(modifiedDocxPath);

        // -----------------------------------------------------------------
        // Load the two documents for comparison.
        // -----------------------------------------------------------------
        Document docToCompare = new Document(originalDocPath);
        Document docToCompareAgainst = new Document(modifiedDocxPath);

        // Ensure both documents have no revisions before the comparison.
        if (docToCompare.Revisions.Count != 0 || docToCompareAgainst.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");
        }

        // Perform the comparison. The revisions will be added to docToCompare.
        docToCompare.Compare(docToCompareAgainst, "Comparer", DateTime.Now);

        // Save the document that now contains the revision markup.
        docToCompare.Save(comparisonResultPath);

        // -----------------------------------------------------------------
        // Inspect the revisions collection to verify that table differences were detected.
        // -----------------------------------------------------------------
        int tableRevisionCount = 0;
        foreach (Revision rev in docToCompare.Revisions)
        {
            // Revisions related to tables have a parent node of type Table.
            if (rev.ParentNode != null && rev.ParentNode.NodeType == NodeType.Table)
                tableRevisionCount++;
        }

        // Output the verification result.
        Console.WriteLine($"Total revisions detected: {docToCompare.Revisions.Count}");
        Console.WriteLine($"Table‑related revisions detected: {tableRevisionCount}");

        // Validation: at least one table revision should exist.
        if (tableRevisionCount == 0)
        {
            Console.WriteLine("Error: No table differences were detected.");
        }
        else
        {
            Console.WriteLine("Success: Table differences were detected.");
        }
    }
}
