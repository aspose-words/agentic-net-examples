using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the sample files (stored in the current working directory).
        string originalPath = "original.doc";
        string revisedPath = "revised.docx";
        string resultPath = "comparisonResult.docx";

        // -------------------------------------------------
        // Create the original DOC file with a simple 2x2 table.
        // -------------------------------------------------
        Document originalDoc = new Document();
        DocumentBuilder builderOrig = new DocumentBuilder(originalDoc);

        // Insert a table with two rows and two columns.
        Table tableOrig = builderOrig.StartTable();
        builderOrig.InsertCell();
        builderOrig.Write("Cell 1A");
        builderOrig.InsertCell();
        builderOrig.Write("Cell 1B");
        builderOrig.EndRow();

        builderOrig.InsertCell();
        builderOrig.Write("Cell 2A");
        builderOrig.InsertCell();
        builderOrig.Write("Cell 2B");
        builderOrig.EndRow();
        builderOrig.EndTable();

        // Save as legacy DOC format.
        originalDoc.Save(originalPath, SaveFormat.Doc);

        // -------------------------------------------------
        // Create the revised DOCX file with a modified table.
        // -------------------------------------------------
        Document revisedDoc = new Document();
        DocumentBuilder builderRev = new DocumentBuilder(revisedDoc);

        // Insert a table with the same layout but change one cell's text.
        Table tableRev = builderRev.StartTable();
        builderRev.InsertCell();
        builderRev.Write("Cell 1A"); // unchanged
        builderRev.InsertCell();
        builderRev.Write("Cell 1B - edited"); // changed text
        builderRev.EndRow();

        builderRev.InsertCell();
        builderRev.Write("Cell 2A");
        builderRev.InsertCell();
        builderRev.Write("Cell 2B");
        builderRev.EndRow();
        builderRev.EndTable();

        // Save as DOCX.
        revisedDoc.Save(revisedPath, SaveFormat.Docx);

        // -------------------------------------------------
        // Load the documents from disk (demonstrates load usage).
        // -------------------------------------------------
        Document loadedOriginal = new Document(originalPath);
        Document loadedRevised = new Document(revisedPath);

        // -------------------------------------------------
        // Compare the documents. Revisions will be added to the original document.
        // -------------------------------------------------
        loadedOriginal.Compare(loadedRevised, "Comparer", DateTime.Now);

        // Verify that at least one revision exists.
        if (loadedOriginal.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Count revisions that affect table content.
        int tableRevisionCount = 0;
        foreach (Revision rev in loadedOriginal.Revisions)
        {
            // A revision caused by a change inside a table has a parent node (e.g., Paragraph)
            // whose ancestor is a Table node.
            if (rev.ParentNode != null && rev.ParentNode.GetAncestor(NodeType.Table) != null)
                tableRevisionCount++;
        }

        // Ensure that table differences were detected.
        if (tableRevisionCount == 0)
            throw new InvalidOperationException("No table revisions were detected, but a difference was expected.");

        // Save the comparison result for visual inspection.
        loadedOriginal.Save(resultPath, SaveFormat.Docx);
    }
}
