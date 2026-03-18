using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocuments
{
    static void Main()
    {
        // Load the two versions of the document.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Both documents must be revision‑free before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Perform the comparison. Revisions are added to docOriginal.
            docOriginal.Compare(docEdited, "John Doe", DateTime.Now);
        }

        // Output information about each revision (optional).
        foreach (Revision rev in docOriginal.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Author: {rev.Author}, Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the document that now contains the tracked changes.
        docOriginal.Save("RevisionResult.docx");
    }
}
