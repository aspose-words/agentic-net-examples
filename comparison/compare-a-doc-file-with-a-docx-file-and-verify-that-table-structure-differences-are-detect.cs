using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create the original DOC file with a simple 1‑row, 2‑cell table.
        Document originalDoc = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(originalDoc);
        builderOriginal.StartTable();
        builderOriginal.InsertCell();
        builderOriginal.Write("Original Cell 1");
        builderOriginal.InsertCell();
        builderOriginal.Write("Original Cell 2");
        builderOriginal.EndTable();

        const string originalPath = "Original.doc";
        originalDoc.Save(originalPath, SaveFormat.Doc);

        // Create the revised DOCX file with a different table structure (1‑row, 3‑cell).
        Document revisedDoc = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revisedDoc);
        builderRevised.StartTable();
        builderRevised.InsertCell();
        builderRevised.Write("Revised Cell A");
        builderRevised.InsertCell();
        builderRevised.Write("Revised Cell B");
        builderRevised.InsertCell();
        builderRevised.Write("Revised Cell C");
        builderRevised.EndTable();

        const string revisedPath = "Revised.docx";
        revisedDoc.Save(revisedPath, SaveFormat.Docx);

        // Load the two documents for comparison.
        Document docToCompare = new Document(originalPath);
        Document docReference = new Document(revisedPath);

        // Perform the comparison. Revisions will be added to docToCompare.
        docToCompare.Compare(docReference, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (docToCompare.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Verify that a revision affecting a table was detected.
        bool tableRevisionFound = false;
        foreach (Revision rev in docToCompare.Revisions)
        {
            // Some table‑related revisions have a Paragraph as the ParentNode.
            // Check whether the revision node is inside a Table by walking up the ancestor chain.
            if (rev.ParentNode != null && rev.ParentNode.GetAncestor(NodeType.Table) != null)
            {
                tableRevisionFound = true;
                break;
            }
        }

        if (!tableRevisionFound)
            throw new InvalidOperationException("Expected a table revision but none was found.");

        // Save the comparison result.
        const string resultPath = "ComparisonResult.docx";
        docToCompare.Save(resultPath, SaveFormat.Docx);
    }
}
